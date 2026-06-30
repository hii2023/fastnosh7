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
  pcod:       { monthly: 4999, trial: 1250 },
  fruit:      { monthly: 4999, trial: 1250 },
};
const UNITS = { monthly: 25, trial: 5 };
const ADDON_PRICE = { fruit: 149, protein: 80, drink: 49 }; // per meal
const PROMOS = { HEALTHY: { monthly: 100, trial: 150 } }; // rupees off per plan (must match index.html PROMOS)
// Distance-fee constants (must match index.html CONFIG).
const BASE_LAT = 23.0299, BASE_LNG = 72.5119;
const FREE_KM_LIMIT = 5.2, BASE_KM = 5, PER_KM_FEE = 10, ROAD_FACTOR = 1.3;

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
        const base = CATPRICE[cat][body.plan];
        if (!base) return json({ error: "invalid plan" }, 400);
        // add-ons: trust only WHICH ones were picked, price them ourselves (per meal x units)
        let addonPerMeal = 0;
        if (Array.isArray(body.addons)) {
          for (const k of body.addons) { if (ADDON_PRICE[k]) addonPerMeal += ADDON_PRICE[k]; }
        }
        const feePerDelivery = distanceFeePerDelivery(body.lat, body.lng);
        // promo: trust only the code, apply our own per-plan discount
        const code = (body.promo || "").toString().trim().toUpperCase();
        const discount = (code && PROMOS[code]) ? (PROMOS[code][body.plan] || 0) : 0;
        const rupees = Math.max(1, base + addonPerMeal * units + feePerDelivery * units - discount);
        if (!rupees || rupees < 1) return json({ error: "invalid amount" }, 400);

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
