# Nosh7 Razorpay backend

A tiny Cloudflare Worker that makes Razorpay payments trusted. It does the two things
that must never happen in the browser:

1. **Creates the order** using your Razorpay secret (so the amount cannot be tampered with).
2. **Verifies the signature** after payment (so a fake success cannot be faked).

## Deploy (one time, about 5 minutes)

You need a Razorpay account and the Cloudflare CLI (`npm i -g wrangler`).

1. Get your keys: Razorpay Dashboard, Settings, API Keys. Copy the **Key Id** and **Key Secret**.
   Use Test Mode keys first (`rzp_test_...`), switch to Live keys when ready.

2. From this folder, log in and set the secrets:

   ```sh
   cd razorpay-worker
   wrangler login
   wrangler secret put RAZORPAY_KEY_ID
   wrangler secret put RAZORPAY_KEY_SECRET
   ```

3. Deploy:

   ```sh
   wrangler deploy
   ```

   Wrangler prints a URL like `https://nosh7-pay.<your-subdomain>.workers.dev`.

4. Open `../index.html` and fill the CONFIG block at the top of the script:

   ```js
   PAYMENT_METHOD: "razorpay",
   RAZORPAY_KEY_ID: "rzp_test_xxxxxxxx",   // the Key Id only (the secret stays in the worker)
   RAZORPAY_ORDER_ENDPOINT:  "https://nosh7-pay.<your-subdomain>.workers.dev/create-order",
   RAZORPAY_VERIFY_ENDPOINT: "https://nosh7-pay.<your-subdomain>.workers.dev/verify",
   ```

5. Lock down CORS: edit `wrangler.toml`, uncomment `ALLOW_ORIGIN` and set it to the site origin
   (for example `https://order.nosh7.in`), then `wrangler deploy` again.

## Test

With Test Mode keys, Razorpay shows test cards (for example `4111 1111 1111 1111`, any future
expiry, any CVV). A successful test payment lands on the confirmation screen only after the worker
verifies the signature.

## Notes

- The Key **Secret** lives only inside the worker. The page only ever holds the public Key Id.
- If you leave the endpoints blank, the page falls back to the UPI deep link, so nothing breaks.
- Add a `/webhook` route later if you want Razorpay to notify you of captures out of band.
