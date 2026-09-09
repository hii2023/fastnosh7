// n7-ingest — receives start.nosh7.in orders and persists them via n7_ingest_funnel_order,
// then best-effort sends a WhatsApp order-confirmation template (if configured + connected).
//
// A "paid" order is auto-activated ONLY when the payment is proven genuine. Proof is preferred
// straight from Razorpay (GET /v1/payments/{id} -> status "captured"), which is the source of
// truth and cannot silently fail the way the old shared-secret ticket did. The signed ticket is
// kept as a fallback for when Razorpay is briefly unreachable.
import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const TICKET_SECRET = Deno.env.get("ORDER_TICKET_SECRET") ?? "";
const RZP_KEY_ID = Deno.env.get("RAZORPAY_KEY_ID") ?? "";
const RZP_KEY_SECRET = Deno.env.get("RAZORPAY_KEY_SECRET") ?? "";
const ADMIN_KEY = Deno.env.get("ADMIN_INGEST_KEY") ?? "";
const GRAPH = "https://graph.facebook.com/v22.0";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "content-type, x-admin-key",
};
const json = (b: unknown, s = 200) =>
  new Response(JSON.stringify(b), { status: s, headers: { ...CORS, "Content-Type": "application/json" } });

async function sb(path: string, init: RequestInit = {}) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...init,
    headers: { apikey: SERVICE_KEY, Authorization: `Bearer ${SERVICE_KEY}`, "Content-Type": "application/json",
               Prefer: "return=representation", ...(init.headers || {}) },
  });
  const t = await r.text();
  return { ok: r.ok, data: t ? JSON.parse(t) : null };
}

async function hmacHex(secret: string, message: string): Promise<string> {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const sig = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(message));
  return [...new Uint8Array(sig)].map((b) => b.toString(16).padStart(2, "0")).join("");
}
function safeEqual(a: string, b: string) {
  if (a.length !== b.length) return false;
  let r = 0; for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}
function digits(p: string) { return String(p || "").replace(/\D/g, ""); }

// Read back the FINALIZED order from the DB (service key, no RLS) so a caller can mirror the full
// details — crucially the delivery ADDRESS, which lives only here (Razorpay never stores it) —
// into the Excel sheet. Returns a flat object the worker maps to sheet columns, or null.
async function finalizedOrder(orderNo: string): Promise<Record<string, any> | null> {
  if (!orderNo) return null;
  try {
    const sel = "id,customer_id,source_ref,status,plan_label,units_total,slot,diet,delivery_days," +
      "start_date,price_rupees,dist_fee_per_delivery,delivery_fee_total,distance_km,address," +
      "lat,lng,instructions,addons";
    const q = await sb(`n7_subscriptions?source_ref=eq.${encodeURIComponent(orderNo)}&select=${sel}` +
      `&order=created_at.desc&limit=1`);
    const s = q.data && q.data[0];
    if (!s) return null;
    let name = "", phone = "", house = "", building = "", area = "", pincode = "";
    if (s.customer_id) {
      const c = await sb(`n7_customers?id=eq.${s.customer_id}&select=name,phone,house,building,area,pincode`);
      const cu = c.data && c.data[0];
      if (cu) { name = cu.name || ""; phone = cu.phone || ""; house = cu.house || "";
        building = cu.building || ""; area = cu.area || ""; pincode = cu.pincode || ""; }
    }
    const instr = Array.isArray(s.instructions) ? s.instructions.filter(Boolean).join(", ")
                : (s.instructions || "");
    return {
      name, phone, house, building, area, pincode,
      plan_label: s.plan_label || "", units_total: s.units_total ?? "",
      slot: s.slot || "", diet: s.diet || "",
      delivery_days: Array.isArray(s.delivery_days) ? s.delivery_days : [],
      start_date: s.start_date || "", price_rupees: s.price_rupees ?? "",
      dist_fee_per_delivery: s.dist_fee_per_delivery ?? "", delivery_fee_total: s.delivery_fee_total ?? "",
      distance_km: s.distance_km ?? "", address: s.address || "",
      lat: s.lat ?? "", lng: s.lng ?? "", instructions: instr,
      status: s.status || "",
    };
  } catch (_e) { return null; }
}

// --- Razorpay source-of-truth lookups (uses the same live key pair as the worker) ---
async function rzpGet(path: string): Promise<any | null> {
  if (!RZP_KEY_ID || !RZP_KEY_SECRET) return null;
  try {
    const auth = "Basic " + btoa(`${RZP_KEY_ID}:${RZP_KEY_SECRET}`);
    const r = await fetch("https://api.razorpay.com" + path, { headers: { Authorization: auth } });
    if (!r.ok) return null;
    return await r.json();
  } catch (_e) { return null; }
}
// Returns the payment entity (or null).
async function razorpayPayment(paymentId: string): Promise<any | null> {
  if (!paymentId) return null;
  const p = await rzpGet(`/v1/payments/${encodeURIComponent(paymentId)}`);
  return p && p.id ? p : null;
}

// --- Label -> structured key helpers (recover a lost order from Razorpay notes) ---
function planFromText(t: string): string {
  const s = String(t || "").toLowerCase();
  const two = /2\s*meal|two\s*meal|breakfast\s*&|lunch\s*&/.test(s);
  if (/trial|weekly|trial pack/.test(s)) return two ? "trial2" : "trial";
  if (/month/.test(s)) return two ? "monthly2" : "monthly";
  return "";
}
function catFromText(t: string): string {
  const s = String(t || "").toLowerCase();
  if (/high protein|muscle|protein/.test(s)) return "protein";
  if (/low ?sugar/.test(s)) return "lowsugar";
  if (/weight ?loss/.test(s)) return "weightloss";
  if (/vegan|plant/.test(s)) return "vegan";
  if (/fruit/.test(s)) return "fruit";
  if (/fresh|healthy/.test(s)) return "fresh";
  return "";
}
function dietFromText(t: string): string {
  const s = String(t || "").toLowerCase();
  if (/swami/.test(s)) return "swami";
  if (/jain/.test(s)) return "jain";
  return "regular";
}
function daysFromText(t: string): string[] {
  const s = String(t || "").toLowerCase();
  const WK = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
  if (/mon\s*to\s*fri/.test(s)) return ["mon", "tue", "wed", "thu", "fri"];
  if (/mon\s*to\s*sat/.test(s)) return ["mon", "tue", "wed", "thu", "fri", "sat"];
  const found = WK.filter((d) => s.includes(d));
  return found.length ? found : ["mon", "tue", "wed", "thu", "fri", "sat"];
}
const UNITS: Record<string, number> = { trial: 5, monthly: 25, trial2: 10, monthly2: 50 };

// Rebuild a full structured order payload from a captured Razorpay payment + its order.
// Prefers structured keys stamped on the order notes by the worker (going forward); falls
// back to the human labels the checkout put on the payment notes (historical orders).
async function reconstructFromRazorpay(orderNo: string, pay: any): Promise<Record<string, any>> {
  const pn = pay.notes || {};
  let on: Record<string, any> = {};
  if (pay.order_id) {
    const od = await rzpGet(`/v1/orders/${encodeURIComponent(pay.order_id)}`);
    if (od && od.notes) on = od.notes;
    if (!orderNo && od && od.receipt) orderNo = String(od.receipt);
  }
  const plan = String(on.plan || "") || planFromText(pn.plan || on.plan_label || "");
  const category = String(on.cat || on.category || "") || catFromText(pn.plan || "");
  const units = Number(on.units || 0) || UNITS[plan] || 0;
  const diet = String(on.diet || "") || dietFromText(pn.diet || "");
  const slot = String(on.slot || pn.slot || "");
  const days = on.days ? String(on.days).split(",").map((x: string) => x.trim()).filter(Boolean)
                       : daysFromText(pn.days || "");
  const addons = on.addons ? String(on.addons).split(",").map((x: string) => x.trim()).filter(Boolean) : [];
  const addrParts = [on.house, on.building, on.area].map((x) => String(x || "").trim()).filter(Boolean);
  return {
    orderNo,
    phone: String(on.phone || pn.phone || pay.contact || ""),
    name: String(on.name || pn.name || ""),
    status: "paid",
    paymentId: pay.id,
    plan, category, deliveries: units,
    diet: diet || "regular", slot,
    days,
    addons,
    house: on.house || "", building: on.building || "", area: on.area || "", pincode: on.pin || on.pincode || "",
    address: addrParts.join(", "),
    total: Math.round((pay.amount || 0) / 100),
    instructions: String(pn.instructions || ""),
    _reconstructed: true,
  };
}

async function ingest(payload: Record<string, any>, verified: boolean) {
  const resp = await fetch(`${SUPABASE_URL}/rest/v1/rpc/n7_ingest_funnel_order`, {
    method: "POST",
    headers: { "Content-Type": "application/json", apikey: SERVICE_KEY, Authorization: `Bearer ${SERVICE_KEY}` },
    body: JSON.stringify({ p: payload, p_verified: verified }),
  });
  const text = await resp.text();
  let result: any; try { result = JSON.parse(text); } catch { result = text; }
  return { ok: resp.ok, result, text };
}

async function maybeConfirm(payload: Record<string, any>, result: any) {
  try {
    if (!result?.ok || !(result.verified && result.status === "active")) return;
    const cfg = await sb(`n7_config?key=in.(wa_confirm_enabled,wa_confirm_template,wa_confirm_lang,wa_confirm_vars)&select=key,value`);
    const map: Record<string, string> = {};
    for (const row of cfg.data || []) map[row.key] = row.value;
    if (map.wa_confirm_enabled !== "1" || !map.wa_confirm_template) return;

    const st = await sb(`wa_settings?id=eq.1&select=access_token,phone_number_id,connected`);
    const s = st.data?.[0];
    if (!s || !s.connected || !s.access_token || !s.phone_number_id) return;

    const to = digits(payload.phone);
    if (!to) return;

    const fields = (map.wa_confirm_vars || "").split(",").map((x) => x.trim()).filter(Boolean);
    const val = (f: string) =>
      f === "name" ? (payload.name || "") :
      f === "plan" ? (payload.planLabel || payload.plan || "") :
      f === "units" ? String(payload.deliveries ?? "") :
      f === "order" ? (payload.orderNo || "") :
      f === "start" ? (payload.startDate || "") : "";
    const components = fields.length
      ? [{ type: "body", parameters: fields.map((f) => ({ type: "text", text: String(val(f) || " ") })) }]
      : [];

    const r = await fetch(`${GRAPH}/${s.phone_number_id}/messages`, {
      method: "POST", headers: { Authorization: `Bearer ${s.access_token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ messaging_product: "whatsapp", to, type: "template",
        template: { name: map.wa_confirm_template, language: { code: map.wa_confirm_lang || "en_US" },
          ...(components.length ? { components } : {}) } }),
    });
    const d = await r.json();
    await sb(`n7_notifications`, { method: "POST", body: JSON.stringify({
      customer_id: result.customer_id, channel: "whatsapp", kind: "order_confirmation",
      status: r.ok ? "sent" : "failed", error: r.ok ? null : d?.error?.message,
      sent_at: r.ok ? new Date().toISOString() : null,
      payload: { orderNo: payload.orderNo, wa_message_id: d?.messages?.[0]?.id ?? null } }) });
  } catch (_e) { /* never break ingestion */ }
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return json({ ok: false, error: "method_not_allowed" }, 405);

  let payload: Record<string, any>;
  try { payload = JSON.parse(await req.text()); } catch { return json({ ok: false, error: "bad_json" }, 400); }

  // --- Admin diagnostic / recovery (gated by ADMIN_INGEST_KEY) ---
  // action:"peek"      -> return what Razorpay knows about a payment (no DB write)
  // action:"reconcile" -> rebuild the full order from Razorpay and re-ingest it as paid/active
  const adminKey = req.headers.get("x-admin-key") || String(payload.adminKey || "");
  const action = String(payload.action || "");
  if (action === "peek" || action === "reconcile") {
    if (!ADMIN_KEY || !safeEqual(adminKey, ADMIN_KEY)) return json({ ok: false, error: "unauthorized" }, 401);
    const pay = await razorpayPayment(String(payload.paymentId || ""));
    if (!pay) return json({ ok: false, error: "payment_not_found" }, 404);
    const captured = pay.status === "captured";
    if (action === "peek") {
      const built = await reconstructFromRazorpay(String(payload.orderNo || ""), pay);
      return json({ ok: true, captured, status: pay.status, amount: pay.amount, notes: pay.notes, order_id: pay.order_id, reconstructed: built });
    }
    // reconcile
    if (!captured) return json({ ok: false, error: "not_captured", status: pay.status }, 409);
    const built = await reconstructFromRazorpay(String(payload.orderNo || ""), pay);
    // allow the caller to patch in fields Razorpay can't know (e.g. address of a lost order)
    if (payload.override && typeof payload.override === "object") Object.assign(built, payload.override);
    const out = await ingest(built, true);
    if (!out.ok) return json({ ok: false, error: "rpc_failed", detail: out.text }, 502);
    await maybeConfirm(built, out.result);
    const order = await finalizedOrder(String(built.orderNo || payload.orderNo || ""));
    return json({ ok: true, reconciled: true, result: out.result, used: built, order });
  }

  // --- Normal funnel/webhook ingestion ---
  const orderNo = String(payload.orderNo ?? "");
  const phone = String(payload.phone ?? "");
  if (!orderNo || !phone) return json({ ok: false, error: "missing_orderNo_or_phone" }, 400);

  let verified = false;
  const status = String(payload.status ?? "pending").toLowerCase();
  const ticket = String(payload.ticket ?? "");
  const paymentId = String(payload.paymentId ?? "");
  let payStatus = "";
  let pay: any = null;
  if (status === "paid" && paymentId) {
    // 1) Source of truth: ask Razorpay whether this payment is actually captured.
    pay = await razorpayPayment(paymentId);
    if (pay) { payStatus = pay.status; if (pay.status === "captured") verified = true; }
    // 2) Fallback: the signed ticket (only if Razorpay was unreachable AND the secret is set).
    if (!verified && !pay && ticket && TICKET_SECRET) {
      verified = safeEqual(await hmacHex(TICKET_SECRET, `${orderNo}|${paymentId}`), ticket);
    }
  }

  // Self-heal a thin paid order: if the browser's full pre-write was lost (or only the
  // minimal webhook reached us), rebuild the missing order details straight from Razorpay
  // so it activates with the RIGHT plan/units/diet/slot instead of a fallback guess.
  // Only FILLS BLANKS — never overwrites a value the payload already carries.
  if (verified && pay && (!payload.plan || !(Number(payload.deliveries) > 0))) {
    const built = await reconstructFromRazorpay(orderNo, pay);
    for (const k of Object.keys(built)) {
      const cur = (payload as any)[k];
      const empty = cur === undefined || cur === null || cur === "" || (Array.isArray(cur) && cur.length === 0);
      const v = (built as any)[k];
      const hasV = v !== undefined && v !== null && v !== "" && !(Array.isArray(v) && v.length === 0);
      if (empty && hasV) (payload as any)[k] = v;
    }
  }

  const out = await ingest(payload, verified);
  if (!out.ok) return json({ ok: false, error: "rpc_failed", detail: out.text }, 502);

  await maybeConfirm(payload, out.result);
  // Return the finalized portal row so the worker's webhook can mirror the full details
  // (esp. the ADDRESS, which lives only in the portal) into the Excel sheet.
  const order = await finalizedOrder(orderNo);
  return json({ ok: true, verified, payStatus, result: out.result, order });
});
