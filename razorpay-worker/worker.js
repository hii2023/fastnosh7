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
        const rupees = Number(body.amount);
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
