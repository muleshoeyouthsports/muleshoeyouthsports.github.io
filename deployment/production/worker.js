const origin = "https://muleshoeyouthsports.github.io";
const headers = {
  "Content-Type": "application/json",
  "Access-Control-Allow-Origin": origin,
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  "Cache-Control": "no-store",
  Vary: "Origin"
};
const reply = (data, status = 200) => new Response(JSON.stringify(data), { status, headers });
const fail = (message, status) => reply({ success: false, message }, status);
const uuid = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i;
const tokenPattern = /^[A-Za-z0-9_-]{43}$/;

export default {
  async fetch(request, env) {
    const path = new URL(request.url).pathname;
    if (path === "/stripe-webhook") return webhook(request, env);
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers });
    if (path === "/" && request.method === "GET") {
      return reply({ application: "MYSA Basketball Registration", registrationEnabled: true });
    }
    if (request.method !== "POST") return fail("Method not allowed.", 405);
    if (path === "/connection-test") {
      if (!env.MYSA_TEST_KEY || request.headers.get("X-MYSA-Test-Key") !== env.MYSA_TEST_KEY) {
        return fail("Unauthorized.", 403);
      }
      try {
        const result = await fetch(env.GOOGLE_SCRIPT_URL);
        return reply({ success: result.ok, googleStatus: result.status, googleResponse: await result.json() },
          result.ok ? 200 : 502);
      } catch {
        return fail("Google connection failed.", 502);
      }
    }
    if (path !== "/" && path !== "/create-checkout-session") return fail("Not found.", 404);
    if (request.headers.get("Origin") !== origin) return fail("Invalid origin.", 403);
    let body;
    try {
      body = await request.json();
    } catch {
      return fail("Invalid JSON.", 400);
    }
    if (!body || typeof body !== "object" || Array.isArray(body)) return fail("Invalid request.", 400);
    try {
      if (path === "/create-checkout-session") {
        if (env.MYSA_LIVE_CARDS !== "true") return fail("Card payments are not open yet.", 503);
        if (typeof body.familyId !== "string" || !uuid.test(body.familyId) ||
            typeof body.token !== "string" || !tokenPattern.test(body.token)) {
          return fail("Invalid payment reference.", 400);
        }
        return reply(await appsScript(env, {
          action: "paymentCheckout", familyId: body.familyId, token: body.token
        }));
      }

      const { turnstileToken, action, ...registration } = body;
      if (registration.paymentMethod === "Stripe" && env.MYSA_LIVE_CARDS !== "true") {
        return fail("Card payments are not open yet.", 503);
      }
      if (typeof turnstileToken !== "string" || !turnstileToken || turnstileToken.length > 2048) {
        return fail("Please complete security verification.", 400);
      }
      const response = await fetch("https://challenges.cloudflare.com/turnstile/v0/siteverify", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ secret: env.TURNSTILE_SECRET_KEY, response: turnstileToken })
      });
      const result = await response.json();
      if (!response.ok || result.success !== true || result.hostname !== new URL(origin).hostname) {
        return fail("Security verification failed.", 403);
      }
      return reply(await appsScript(env, registration));
    } catch {
      return fail("Registration service is temporarily unavailable.", 502);
    }
  }
};

async function appsScript(env, payload) {
  if (!env.GOOGLE_SCRIPT_URL || !env.MYSA_API_KEY) throw new Error("Apps Script is not configured");
  const response = await fetch(env.GOOGLE_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ ...payload, apiKey: env.MYSA_API_KEY })
  });
  if (!response.ok) throw new Error("Apps Script request failed");
  return response.json();
}

async function webhook(request, env) {
  if (request.method !== "POST") return new Response("Method not allowed", { status: 405 });
  if (!env.STRIPE_LIVE_WEBHOOK_SECRET) return new Response("Webhook not configured", { status: 503 });
  const raw = await request.text();
  if (!(await validSignature(raw, request.headers.get("Stripe-Signature"), env.STRIPE_LIVE_WEBHOOK_SECRET))) {
    return new Response("Invalid signature", { status: 400 });
  }
  let event;
  try { event = JSON.parse(raw); } catch { return new Response("Invalid event", { status: 400 }); }
  const session = event?.data?.object;
  if (event?.livemode !== true ||
      !["checkout.session.completed", "checkout.session.async_payment_succeeded"].includes(event.type) ||
      session?.object !== "checkout.session" || session.payment_status !== "paid" ||
      typeof session.id !== "string" || !/^cs_live_[A-Za-z0-9]+$/.test(session.id) ||
      typeof session.client_reference_id !== "string" || !uuid.test(session.client_reference_id) ||
      session.metadata?.family_id !== session.client_reference_id ||
      session.metadata?.purpose !== "mysa-basketball-registration") {
    return new Response(null, { status: 204 });
  }
  try {
    const result = await appsScript(env, {
      action: "paymentFulfill", familyId: session.client_reference_id, sessionId: session.id
    });
    if (result.success === true) return new Response(null, { status: 204 });
  } catch {
    // Stripe retries when fulfillment fails.
  }
  return new Response("Payment update failed", { status: 502 });
}

async function validSignature(raw, header, secret) {
  const parts = (header || "").split(",").map(part => part.trim());
  const timestamp = parts.find(part => part.startsWith("t="))?.slice(2);
  if (!timestamp || !/^\d+$/.test(timestamp) ||
      Math.abs(Date.now() / 1000 - Number(timestamp)) > 300) return false;
  const encoder = new TextEncoder();
  const key = await crypto.subtle.importKey("raw", encoder.encode(secret),
    { name: "HMAC", hash: "SHA-256" }, false, ["verify"]);
  const message = encoder.encode(timestamp + "." + raw);
  for (const part of parts.filter(part => /^v1=[0-9a-f]{64}$/i.test(part))) {
    const signature = Uint8Array.from(part.slice(3).match(/.{2}/g), hex => parseInt(hex, 16));
    if (await crypto.subtle.verify("HMAC", key, signature, message)) return true;
  }
  return false;
}
