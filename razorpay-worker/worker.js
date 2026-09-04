/**
 * Nosh7 Razorpay backend (Cloudflare Worker)
 *
 * Two routes:
 *   POST /create-order  ->  creates a Razorpay order, returns { id, amount, currency }
 *   POST /verify        ->  verifies the payment signature, returns { valid: true|false }
 *
 * Secrets (set with `wrangler secret put`):
 *   RAZORPAY_KEY_ID      (rzp_live_xxx or rzp_test_xxx)
 *   RAZORPAY_KEY_SECRET  (the matching secret, never shipped to the browser)
 *
 * Optional var (wrangler.toml [vars]):
 *   ALLOW_ORIGIN         (the site origin, e.g. https://order.nosh7.in). Defaults to "*".
 */

// Authoritative price tables (must match the front-end CATS / ADDONS).
// Price depends on the category + plan. Standard tracks share one price.
const CATPRICE = {
  fresh:      { monthly: 4999, trial: 1250 },
  protein:    { monthly: 6999, trial: 1650 },
  lowsugar:   { monthly: 4999, trial: 1250 },
  weightloss: { monthly: 5999, trial: 1445 },
  vegan:      { monthly: 5975, trial: 1445 },
  fruit:      { monthly: 5250, trial: 1250 },
};
const UNITS = { monthly: 25, trial: 5, monthly2: 50, trial2: 10 };
// 2-meals-a-day "Daily Plan": maps to its 1-meal base plan; base price = 2x, minus DAILY_DISCOUNT.
// (must match index.html plansFor / DAILY_DISCOUNT)
const DAILY = { monthly2: "monthly", trial2: "trial" };
const DAILY_DISCOUNT = 0.07;
const ADDON_PRICE = { fruit: 169, protein: 80, drink: 49 }; // per meal
// Promos are managed in the portal and applied by n7_quote (the authoritative quote below).
// The local price tables here are ONLY a fallback for when the portal quote is unreachable, and
// in that same outage the funnel is fail-closed (offers no promo), so the fallback applies NO
// promo discount too — the two must never disagree on what the customer was shown vs charged.
// Distance-fee constants (fallback if the portal quote is unreachable; must match index.html CONFIG).
const BASE_LAT = 23.0299, BASE_LNG = 72.5119;
const FREE_KM_LIMIT = 5.2, BASE_KM = 5, PER_KM_FEE = 10, ROAD_FACTOR = 1.3;

// Portal quote (source of truth for delivery charge + promo, configurable from the admin Settings).
// Public RPC, publishable/anon key (safe in a server; exposes no customer data).
const SB_URL = "https://xoiksbtxoxrifkgvupqp.supabase.co";
const SB_ANON = "sb_publishable_UwUNp74KFKUlbqQ7aY0s7Q_toy7kUiD";
// Ask the portal for the authoritative payable (Rs). Returns null on any problem so the
// caller falls back to the local price tables below (payments never depend on the portal).
async function portalQuoteRupees(body) {
  try {
    const r = await fetch(SB_URL + "/rest/v1/rpc/n7_quote", {
      method: "POST",
      headers: { "Content-Type": "application/json", apikey: SB_ANON, Authorization: "Bearer " + SB_ANON },
      body: JSON.stringify({ p: {
        category: body.category || "", plan: body.plan,
        addons: Array.isArray(body.addons) ? body.addons : [],
        lat: body.lat, lng: body.lng, promo: body.promo || "", express: body.express === true,
      } }),
    });
    if (!r.ok) return null;
    const q = await r.json();
    if (q && q.ok === true && Number(q.total) >= 1) return Math.round(Number(q.total));
    return null;
  } catch (_e) { return null; }
}

// Portal backfill safety net. On a captured payment we tell the portal (same n7-ingest edge fn
// the browser posts to) that the order is paid, carrying the SAME signed ticket the browser path
// would, so n7-ingest verifies it and activates the subscription. This closes the gap where a
// dropped post-payment beacon left the order in Excel but missing from the portal. Idempotent:
// n7_ingest_funnel_order flips the pre-written pending order to paid and dedups the payment/
// ledger/deliveries, so redelivered webhooks are safe. Returns true on a 2xx.
async function portalIngestPaid(env, o) {
  try {
    const ticket = env.ORDER_TICKET_SECRET
      ? await hmacHex(env.ORDER_TICKET_SECRET, `${o.orderNo}|${o.paymentId || ""}`)
      : "";
    const r = await fetch(SB_URL + "/functions/v1/n7-ingest", {
      method: "POST",
      headers: { "Content-Type": "application/json", apikey: SB_ANON, Authorization: "Bearer " + SB_ANON },
      body: JSON.stringify({
        orderNo: o.orderNo, phone: o.phone || "", name: o.name || "",
        status: "paid", paymentId: o.paymentId || "", ticket, total: o.amountRupees || 0,
      }),
    });
    return r.ok;
  } catch (_e) { return false; }
}

function haversineKm(la1, lo1, la2, lo2) {
  const R = 6371, toR = (x) => (x * Math.PI) / 180;
  const dLa = toR(la2 - la1), dLo = toR(lo2 - lo1);
  const a = Math.sin(dLa / 2) ** 2 + Math.cos(toR(la1)) * Math.cos(toR(la2)) * Math.sin(dLo / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}
// Rupees per delivery; free up to 5.2 km, else Rs 10 per started km beyond 5. Same logic as the app.
function distanceFeePerDelivery(lat, lng) {
  if (lat == null || lng == null || isNaN(Number(lat)) || isNaN(Number(lng))) return 0;
  const km = haversineKm(BASE_LAT, BASE_LNG, Number(lat), Number(lng)) * ROAD_FACTOR;
  return (km <= FREE_KM_LIMIT) ? 0 : Math.max(0, Math.ceil(km - BASE_KM - 1e-9)) * PER_KM_FEE;
}

export default {
  async fetch(request, env) {
    const origin = env.ALLOW_ORIGIN || "*";
    const cors = {
      "Access-Control-Allow-Origin": origin,
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    };

    if (request.method === "OPTIONS") return new Response(null, { headers: cors });

    const url = new URL(request.url);
    const json = (obj, status = 200) =>
      new Response(JSON.stringify(obj), { status, headers: { ...cors, "Content-Type": "application/json" } });

    try {
      if (url.pathname === "/create-order" && request.method === "POST") {
        const body = await request.json();

        // SERVER IS THE SOURCE OF TRUTH FOR PRICE.
        // We never trust an "amount" sent by the browser (it can be tampered).
        // We recompute from the category + plan + which add-ons were chosen,
        // using our own price tables, plus the distance fee from the coords.
        const units = UNITS[body.plan];
        if (!units) return json({ error: "invalid plan" }, 400);
        // default to a standard track if category is missing (keeps old clients working)
        const cat = CATPRICE[body.category] ? body.category : "fresh";
        // a Daily (2-meals-a-day) plan prices as 2x its base plan, minus the daily discount
        const basePlan = DAILY[body.plan] || body.plan;
        const catBase = CATPRICE[cat][basePlan];
        if (!catBase) return json({ error: "invalid plan" }, 400);
        const base = DAILY[body.plan] ? Math.round(2 * catBase * (1 - DAILY_DISCOUNT) / 100) * 100 : catBase;   // 2 meals a day, rounded to nearest Rs 100 (must match index.html d2)
        if (!base) return json({ error: "invalid plan" }, 400);
        // add-ons: trust only WHICH ones were picked, price them ourselves (per meal x units)
        let addonPerMeal = 0;
        if (Array.isArray(body.addons)) {
          for (const k of body.addons) { if (ADDON_PRICE[k]) addonPerMeal += ADDON_PRICE[k]; }
        }
        // Express monthly link (?express=...) takes only name + phone, no address, so it
        // carries no distance fee (flat plan price). For every other flow coords are
        // mandatory: without them the distance fee cannot be charged, and a tampered client
        // could omit them to dodge the fee entirely.
        const isExpress = body.express === true;
        if (!isExpress && (body.lat == null || body.lng == null || isNaN(Number(body.lat)) || isNaN(Number(body.lng))))
          return json({ error: "missing delivery location" }, 400);
        const feePerDelivery = isExpress ? 0 : distanceFeePerDelivery(body.lat, body.lng);
        // promo: the portal quote below is the source of truth for discounts. The local
        // fallback intentionally applies NO promo (see PROMOS note above) so it can never
        // undercharge relative to a fail-closed funnel during a portal outage.
        const localRupees = Math.max(1, base + addonPerMeal * units + feePerDelivery * units);
        if (!localRupees || localRupees < 1) return json({ error: "invalid amount" }, 400);
        // Prefer the portal's configurable quote (delivery charge + promo managed in admin Settings);
        // fall back to the local price tables above if the portal is unreachable.
        const portalRupees = await portalQuoteRupees(body);
        const rupees = (portalRupees != null) ? portalRupees : localRupees;

        const auth = "Basic " + btoa(env.RAZORPAY_KEY_ID + ":" + env.RAZORPAY_KEY_SECRET);
        const res = await fetch("https://api.razorpay.com/v1/orders", {
          method: "POST",
          headers: { Authorization: auth, "Content-Type": "application/json" },
          body: JSON.stringify({
            amount: Math.round(rupees * 100), // paise
            currency: "INR",
            receipt: (body.receipt || "nosh7").slice(0, 40),
            notes: { name: body.name || "", phone: body.phone || "" },
          }),
        });
        const data = await res.json();
        if (!res.ok) return json({ error: data.error?.description || "razorpay error" }, 502);
        return json({ id: data.id, amount: data.amount, currency: data.currency });
      }

      if (url.pathname === "/verify" && request.method === "POST") {
        const b = await request.json();
        const expected = await hmacHex(
          env.RAZORPAY_KEY_SECRET,
          `${b.razorpay_order_id}|${b.razorpay_payment_id}`
        );
        const valid = timingSafeEqual(expected, b.razorpay_signature || "");
        // On a genuinely verified payment, mint a ticket the order sheet can trust.
        // ticket = HMAC(ORDER_TICKET_SECRET, orderNo|paymentId). The secret lives only
        // here and in the Apps Script (never in the browser), so a fake "paid" row
        // submitted straight to the sheet cannot carry a valid ticket.
        let ticket = "";
        if (valid && env.ORDER_TICKET_SECRET) {
          ticket = await hmacHex(env.ORDER_TICKET_SECRET, `${b.orderNo || ""}|${b.razorpay_payment_id || ""}`);
        }
        return json({ valid, ticket });
      }

      if (url.pathname === "/razorpay-webhook" && request.method === "POST") {
        // Server-side safety net. Razorpay calls this directly when a payment is captured,
        // so a paid order reaches the sheet even if the customer's browser never fired its
        // post-payment beacon (tab closed, in-app browser killed, network drop). We verify
        // the webhook signature, then POST a "paid" update to the same Apps Script sheet.
        // The sheet MERGES by Order No, so this only flips status/paymentId on the row the
        // browser pre-wrote (address etc. preserved); if no row exists yet it creates a
        // partial one, which still beats losing the order.
        const raw = await request.text();
        const sig = request.headers.get("X-Razorpay-Signature") || "";
        if (!env.RAZORPAY_WEBHOOK_SECRET) return json({ error: "webhook not configured" }, 503);
        const expected = await hmacHex(env.RAZORPAY_WEBHOOK_SECRET, raw);
        if (!timingSafeEqual(expected, sig)) return json({ error: "bad signature" }, 401);

        let evt;
        try { evt = JSON.parse(raw); } catch (e) { return json({ error: "bad json" }, 400); }

        const pay = evt && evt.payload && evt.payload.payment && evt.payload.payment.entity;
        // Act only on a captured (money-in-hand) payment. Ignore authorized/failed/others.
        if (evt && evt.event === "payment.captured" && pay) {
          let orderNo = (pay.notes && pay.notes.order) ? String(pay.notes.order) : "";
          let phone = pay.contact ? String(pay.contact) : "";
          let name = (pay.notes && pay.notes.name) ? String(pay.notes.name) : "";
          // Read the order to recover the receipt (=orderNo) and the {name,phone} notes we set at
          // create-order time. The order notes carry the exact phone the customer typed on the
          // funnel, which is the most reliable key for the portal customer upsert.
          if (pay.order_id) {
            try {
              const auth = "Basic " + btoa(env.RAZORPAY_KEY_ID + ":" + env.RAZORPAY_KEY_SECRET);
              const or = await fetch("https://api.razorpay.com/v1/orders/" + pay.order_id, { headers: { Authorization: auth } });
              if (or.ok) {
                const od = await or.json();
                if (!orderNo) orderNo = String(od.receipt || "");
                if (od.notes && od.notes.phone) phone = String(od.notes.phone);
                if (od.notes && od.notes.name) name = String(od.notes.name);
              }
            } catch (e) { /* fall through with what we have */ }
          }
          if (orderNo) {
            // Mint the same signed ticket the browser path would, so both sinks mark this
            // "verified" (the webhook signature already proved the payment is genuine).
            const ticket = env.ORDER_TICKET_SECRET
              ? await hmacHex(env.ORDER_TICKET_SECRET, `${orderNo}|${pay.id || ""}`)
              : "";
            // (a) Excel sheet (unchanged): merges by Order No, flips status/paymentId on the
            // row the browser pre-wrote; creates a partial row if none exists yet.
            let sheetOk = true;
            if (env.ORDER_WEBHOOK) {
              sheetOk = false;
              const update = { orderNo, status: "paid", paymentId: pay.id || "", ticket, payment: "razorpay" };
              try {
                const pr = await fetch(env.ORDER_WEBHOOK, {
                  method: "POST",
                  headers: { "Content-Type": "application/json" },
                  body: JSON.stringify(update),
                });
                // Apps Script answers 200 with {ok:true} on success, {ok:false,...} on a caught error.
                const txt = pr.ok ? await pr.text() : "";
                sheetOk = pr.ok && /"ok"\s*:\s*true/.test(txt);
              } catch (e) { sheetOk = false; }
            }
            // (b) Portal (Supabase) backfill — the missing half of this safety net. Flips the
            // pre-written pending subscription to paid/active (or creates a last-resort record).
            const portalOk = await portalIngestPaid(env, {
              orderNo, phone, name, paymentId: pay.id || "",
              amountRupees: Math.round((pay.amount || 0) / 100),
            });
            // Both writes are idempotent, so a failure here MUST NOT be swallowed: return non-2xx
            // so Razorpay redelivers the webhook (~24h). The already-written sink simply no-ops on
            // retry. This is how orders went missing before (money taken, nothing recorded).
            if (!sheetOk || !portalOk) return json({ error: "backfill incomplete; retry" }, 502);
          }
        }
        // 200 for genuinely-ignored events (non-captured) and successful writes, so Razorpay
        // does not hammer retries for events we intentionally drop.
        return json({ ok: true });
      }

      return json({ error: "not found" }, 404);
    } catch (e) {
      return json({ error: "server error" }, 500);
    }
  },
};

async function hmacHex(secret, message) {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sig = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(message));
  return [...new Uint8Array(sig)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let out = 0;
  for (let i = 0; i < a.length; i++) out |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return out === 0;
}
