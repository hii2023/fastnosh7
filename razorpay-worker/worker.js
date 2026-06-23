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

// Authoritative price table (must match the front-end plan buttons).
const PLANS = {
  trial:   { price: 1250, units: 5 },
  monthly: { price: 4999, units: 25 },
};
// Distance-fee constants (must match index.html CONFIG).
const BASE_LAT = 23.0299, BASE_LNG = 72.5119;
const FREE_KM = 5, PER_KM_FEE = 10, ROAD_FACTOR = 1.3;

function haversineKm(la1, lo1, la2, lo2) {
  const R = 6371, toR = (x) => (x * Math.PI) / 180;
  const dLa = toR(la2 - la1), dLo = toR(lo2 - lo1);
  const a = Math.sin(dLa / 2) ** 2 + Math.cos(toR(la1)) * Math.cos(toR(la2)) * Math.sin(dLo / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}
// Rupees per delivery; 0 within the free radius. Same floor rounding as the app.
function distanceFeePerDelivery(lat, lng) {
  if (lat == null || lng == null || isNaN(Number(lat)) || isNaN(Number(lng))) return 0;
  const km = haversineKm(BASE_LAT, BASE_LNG, Number(lat), Number(lng)) * ROAD_FACTOR;
  return Math.max(0, Math.floor(km) - FREE_KM) * PER_KM_FEE;
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
        // We recompute: plan price + distance fee, from the plan id and the
        // delivery coordinates, using the same constants as the front-end.
        const plan = PLANS[body.plan];
        if (!plan) return json({ error: "invalid plan" }, 400);
        const feePerDelivery = distanceFeePerDelivery(body.lat, body.lng);
        const rupees = plan.price + feePerDelivery * plan.units;
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
        return json({ valid });
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
