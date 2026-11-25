function mysaLiveCardAction(data) {
  const settings = PropertiesService.getScriptProperties();
  const stripeKey = settings.getProperty('STRIPE_LIVE_RESTRICTED_KEY') || '';
  const secret = settings.getProperty('MYSA_LIVE_PAYMENT_LINK_SECRET');
  const uuid = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i;
  const registering = data.action === 'cardRegistration';
  const fulfilling = data.action === 'paymentFulfill';
  if (!/^(rk|sk)_live_/.test(stripeKey) || !secret ||
      (!fulfilling && settings.getProperty('MYSA_LIVE_CARDS') !== 'true')) {
    throw new Error('Live card payments are not available.');
  }

  const reference = registering ? data.attemptId : data.familyId;
  if (typeof reference !== 'string' || !uuid.test(reference)) {
    throw new Error('Invalid registration reference.');
  }

  let familyId = reference;
  if (registering) {
    const hex = Utilities.computeHmacSha256Signature(
      'mysa-live-registration:' + reference.toLowerCase(), secret
    ).map(byte => (byte & 255).toString(16).padStart(2, '0')).join('');
    familyId = [hex.slice(0, 8), hex.slice(8, 12), hex.slice(12, 16),
      hex.slice(16, 20), hex.slice(20, 32)].join('-');
  }
  const token = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature('mysa-live-payment:' + familyId, secret)
  ).replace(/=+$/, '');
  if (!registering && !fulfilling && !mysaLiveTokenMatches(data.token, token)) {
    throw new Error('Invalid payment reference.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Player Registrations');
    const expected = ['Sibling Count', 'Fee Per Child', 'Payment Method', 'Payment Status',
      'Amount Received', 'Payment Date', 'Collected By', 'Stripe Payment ID', 'Notes', 'First Family Row'];
    const headers = sheet.getRange(1, 1, 1, 26).getDisplayValues()[0];
    if (headers[2] !== 'Family ID' || expected.some((label, i) => headers[16 + i] !== label)) {
      throw new Error('Unexpected registration columns.');
    }
    const lastRow = sheet.getLastRow();
    const rows = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 26).getValues()
      .map((values, index) => ({ values, number: index + 2 }))
      .filter(row => row.values[2] === familyId);

    if (registering && rows.length === 0) {
      const saved = saveRegistration(data, familyId);
      return { success: true, familyId, token, totalAmount: saved.totalAmount };
    }

    const leaders = rows.filter(row => row.values[25] === 1);
    const fee = rows[0]?.values[17];
    if (leaders.length !== 1 || ![40, 45].includes(fee) || rows.some(row =>
      row.values[16] !== rows.length || row.values[17] !== fee || row.values[18] !== 'Stripe')) {
      throw new Error('Card registration unavailable.');
    }
    const amount = fee * rows.length * 100;
    if (registering) return { success: true, familyId, token, totalAmount: amount / 100 };

    const storageKey = 'mysa-live-checkout:' + familyId;
    const attempt = JSON.parse(settings.getProperty(storageKey) || 'null');
    if (fulfilling) {
      if (!attempt || attempt.sessionId !== data.sessionId) {
        throw new Error('Payment does not match the saved Checkout session.');
      }
      const session = mysaLiveStripeSession(stripeKey, data.sessionId);
      mysaLiveRecordPayment(sheet, rows, session, familyId, amount);
      return { success: true };
    }

    if (rows.every(row => row.values[19] === 'Paid')) return { success: true, status: 'paid' };
    if (attempt?.sessionId) {
      const session = mysaLiveStripeSession(stripeKey, attempt.sessionId);
      if (session.payment_status === 'paid' || session.status === 'complete') {
        return { success: true, status: 'processing' };
      }
      if (session.status === 'open') return { success: true, status: 'open', checkoutUrl: session.url };
      if (session.status !== 'expired') throw new Error('Checkout status requires review.');
    }
    if (rows.some(row => row.values[19] !== 'Pending' || row.values[23])) {
      throw new Error('Payment status requires review.');
    }

    const next = attempt && !attempt.sessionId ? attempt : { key: Utilities.getUuid(), started: Date.now() };
    if (Date.now() - next.started >= 20 * 60 * 60 * 1000) {
      throw new Error('A previous Checkout attempt requires review.');
    }
    settings.setProperty(storageKey, JSON.stringify(next));
    const returnUrl = 'https://muleshoeyouthsports.github.io/';
    const session = mysaLiveStripeSession(stripeKey, '', {
      mode: 'payment',
      'payment_method_types[0]': 'card',
      client_reference_id: familyId,
      'metadata[family_id]': familyId,
      'metadata[purpose]': 'mysa-basketball-registration',
      'adaptive_pricing[enabled]': 'false',
      'line_items[0][quantity]': String(rows.length),
      'line_items[0][price_data][currency]': 'usd',
      'line_items[0][price_data][unit_amount]': String(fee * 100),
      'line_items[0][price_data][product_data][name]': 'MYSA Basketball Registration',
      success_url: returnUrl + '?payment=returned#pay=' + familyId + '.' + token,
      cancel_url: returnUrl + '#pay=' + familyId + '.' + token
    }, next.key);
    if (session.livemode !== true || !/^cs_live_[A-Za-z0-9]+$/.test(session.id)) {
      throw new Error('Expected a live Checkout session.');
    }
    next.sessionId = session.id;
    settings.setProperty(storageKey, JSON.stringify(next));
    return { success: true, status: 'open', checkoutUrl: session.url };
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

function mysaLiveRecordPayment(sheet, rows, session, familyId, amount) {
  const paymentId = session.payment_intent;
  if (session.livemode !== true || session.mode !== 'payment' ||
      session.payment_status !== 'paid' || session.status !== 'complete' ||
      session.client_reference_id !== familyId || session.metadata?.family_id !== familyId ||
      session.metadata?.purpose !== 'mysa-basketball-registration' ||
      session.amount_total !== amount || session.currency !== 'usd' ||
      typeof paymentId !== 'string' || !/^pi_[A-Za-z0-9]+$/.test(paymentId)) {
    throw new Error('Stripe payment could not be verified.');
  }
  if (rows.some(({ values }) => !['Pending', 'Paid'].includes(values[19]) ||
      (values[23] && values[23] !== paymentId) ||
      (values[19] === 'Paid' && values[23] !== paymentId))) {
    throw new Error('Existing payment data requires review.');
  }
  const date = rows.find(row => row.values[19] === 'Paid')?.values[21] || new Date();
  for (const { values, number } of rows) {
    if (values[19] === 'Paid') continue;
    sheet.getRange(number, 20, 1, 6).setValues([[
      'Paid', values[25] === 1 ? amount / 100 : 0, date,
      'Stripe', paymentId, 'Payment verified in Stripe'
    ]]);
  }
}

function mysaLiveStripeSession(key, sessionId, fields, idempotencyKey) {
  const headers = { Authorization: 'Bearer ' + key };
  const options = { headers, muteHttpExceptions: true };
  if (fields) {
    headers['Idempotency-Key'] = idempotencyKey;
    options.method = 'post';
    options.contentType = 'application/x-www-form-urlencoded';
    options.payload = Object.entries(fields).map(([name, value]) =>
      encodeURIComponent(name) + '=' + encodeURIComponent(value)).join('&');
  }
  const url = 'https://api.stripe.com/v1/checkout/sessions' + (sessionId ? '/' + sessionId : '');
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) throw new Error('Stripe Checkout request failed.');
  return JSON.parse(response.getContentText());
}

function mysaLiveTokenMatches(actual, expected) {
  if (typeof actual !== 'string' || actual.length !== expected.length) return false;
  let difference = 0;
  for (let index = 0; index < expected.length; index++) {
    difference |= actual.charCodeAt(index) ^ expected.charCodeAt(index);
  }
  return difference === 0;
}
