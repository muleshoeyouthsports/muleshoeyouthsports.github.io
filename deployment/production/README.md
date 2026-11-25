# Production Stripe integration

Keep `MYSA_LIVE_CARDS` unset in both services and leave the staged website branch unmerged while installing the backend. Use the production Stripe account, production spreadsheet, and `mysa-registration` Worker.

## Apps Script

1. Add `LivePayments.gs` to the production spreadsheet's Apps Script project.
2. In `SaveRegistrations.gs`, replace the opening declaration of `saveRegistration` with:

   ```js
   function saveRegistration(registration) {
     const lock = LockService.getScriptLock();
     lock.waitLock(25000);
     try {
       return mysaSaveRegistrationLocked(registration);
     } finally {
       try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
     }
   }

   function mysaSaveRegistrationLocked(registration, existingFamilyId) {
   ```

   In the renamed function, replace `const familyId = Utilities.getUuid();` with `const familyId = existingFamilyId || Utilities.getUuid();`. Wrap the existing `sendRegistrationEmail({...});` call in `try { ... } catch { console.error('Registration saved, but confirmation email failed. Family ID: ' + familyId); }`. Both paths must use the same script lock for row allocation.
3. In `WebApp.gs`, immediately after the authenticated request passes the API key check and `delete registration.apiKey;`, insert:

   ```js
   if (['paymentCheckout', 'paymentFulfill'].includes(registration.action)) {
     return reply(mysaLiveCardAction(registration));
   }
   if (registration.paymentMethod === 'Stripe') {
     validateOnlineRegistration(registration, true);
     return reply(mysaLiveCardAction({ ...registration, action: 'cardRegistration' }));
   }
   ```

   Change the validator signature to `validateOnlineRegistration(data, allowCard = false)`. Replace its payment-method check with:

   ```js
   if (!['Cash', 'Check'].includes(data.paymentMethod) &&
       !(allowCard && data.paymentMethod === 'Stripe')) {
     throw new Error('Invalid payment method.');
   }
   ```

   Keep the Cash/Check save and JSON response below the new dispatch. Before using the card flow, verify the sheet headers against `LivePayments.gs`.
4. In `RegistrationEmails.gs` change the pending message branch to choose the following wording for Stripe while retaining the existing wording for Cash and Check:

   ```js
   (paymentStatus === 'Pending'
     ? (paymentMethod === 'Stripe'
       ? 'Your registration has been received. Complete your card payment through Stripe Checkout using the registration page.\n\n'
       : 'Your registration has been received. Payment Status: Pending. Please submit your Cash or Check payment to an MYSA board member. Make checks payable to Muleshoe Youth Sports Association.\n\n')
     : 'Your payment has been confirmed.\n\n') +
   ```

5. Add production Apps Script properties `STRIPE_LIVE_RESTRICTED_KEY` (live Checkout Sessions create/read) and `MYSA_LIVE_PAYMENT_LINK_SECRET` (a new random secret). Keep existing `MYSA_API_KEY`. Update the existing web app deployment so its URL is preserved.

## Worker and verification

- Stage `worker.js` on `mysa-registration`. Preserve `GOOGLE_SCRIPT_URL`, `MYSA_API_KEY`, `MYSA_TEST_KEY`, and `TURNSTILE_SECRET_KEY`.
- In live Stripe create a webhook destination for `https://mysa-registration.muleshoeyouthsports.workers.dev/stripe-webhook`, with `checkout.session.completed` and `checkout.session.async_payment_succeeded`. Set the destination's signing secret as Worker secret `STRIPE_LIVE_WEBHOOK_SECRET`.
- Deploy with both card flags unset. Verify the live Cash/Check path and Turnstile before opening cards.
- Turn on `MYSA_LIVE_CARDS=true` in Apps Script and then the Worker for the controlled test. Use a properly authorized live payment, confirm its amount and the matching Paid sheet entry, and check duplicate webhook delivery before merging the website branch.
- To close Checkout again, turn off the Worker card flag. Existing webhook fulfillment continues for payments already in flight.
