/**
 * POST /api/apply — Hydrovac Operator application handler
 *
 * Cloudflare Pages Function. Accepts multipart/form-data, validates,
 * and emails the application to OE with the resume attached via MailChannels.
 *
 * Required environment variables (Cloudflare Pages → Settings → Environment variables):
 *   TO_EMAIL     e.g. opportunities@oeservices.ca
 *   FROM_EMAIL   e.g. careers@oeservices.ca   (must be on a domain you control)
 *   DKIM_DOMAIN      optional — improves deliverability
 *   DKIM_SELECTOR    optional
 *   DKIM_PRIVATE_KEY optional
 */

const MAX_BYTES = 5 * 1024 * 1024;
const ALLOWED_EXT = ['pdf', 'doc', 'docx', 'rtf', 'txt'];

const esc = (s) =>
  String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });

export async function onRequestPost({ request, env }) {
  try {
    const form = await request.formData();

    // Honeypot — silently accept so bots don't retry.
    if (form.get('company_website')) return json({ ok: true });

    const name = (form.get('name') || '').toString().trim();
    const email = (form.get('email') || '').toString().trim();
    const phone = (form.get('phone') || '').toString().trim();
    const licence = (form.get('licence') || '').toString().trim();
    const location = (form.get('location') || '').toString().trim();
    const position = (form.get('position') || 'Hydrovac Operator').toString().trim();
    const message = (form.get('message') || '').toString().trim();
    const resume = form.get('resume');

    if (!name || !email || !phone || !licence || !location) {
      return json({ ok: false, error: 'Missing required fields.' }, 400);
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return json({ ok: false, error: 'Invalid email address.' }, 400);
    }
    if (!resume || typeof resume.arrayBuffer !== 'function') {
      return json({ ok: false, error: 'Resume attachment is required.' }, 400);
    }

    const ext = (resume.name || '').split('.').pop().toLowerCase();
    if (!ALLOWED_EXT.includes(ext)) {
      return json({ ok: false, error: 'Unsupported file type.' }, 400);
    }

    const bytes = await resume.arrayBuffer();
    if (bytes.byteLength > MAX_BYTES) {
      return json({ ok: false, error: 'File exceeds 5 MB.' }, 400);
    }

    // base64 encode in chunks (avoids call-stack overflow on larger files)
    const view = new Uint8Array(bytes);
    let binary = '';
    for (let i = 0; i < view.length; i += 0x8000) {
      binary += String.fromCharCode.apply(null, view.subarray(i, i + 0x8000));
    }
    const b64 = btoa(binary);

    const TO = env.TO_EMAIL || 'opportunities@oeservices.ca';
    const FROM = env.FROM_EMAIL || 'careers@oeservices.ca';

    const html = `
      <div style="font-family:Arial,Helvetica,sans-serif;color:#13203B;line-height:1.6">
        <h2 style="color:#0F2670;margin:0 0 4px">New Application — ${esc(position)}</h2>
        <p style="color:#54607A;margin:0 0 20px;font-size:14px">Submitted via careers.oeservices.ca</p>
        <table cellpadding="7" cellspacing="0" style="border-collapse:collapse;font-size:15px">
          <tr><td style="color:#54607A">Name</td><td><strong>${esc(name)}</strong></td></tr>
          <tr><td style="color:#54607A">Email</td><td><a href="mailto:${esc(email)}">${esc(email)}</a></td></tr>
          <tr><td style="color:#54607A">Phone</td><td>${esc(phone)}</td></tr>
          <tr><td style="color:#54607A">Licence</td><td>${esc(licence)}</td></tr>
          <tr><td style="color:#54607A">Location</td><td>${esc(location)}</td></tr>
        </table>
        ${
          message
            ? `<p style="margin-top:22px"><strong>Why should we hire you?</strong><br>${esc(message).replace(/\n/g, '<br>')}</p>`
            : ''
        }
        <p style="margin-top:22px;font-size:14px;color:#54607A">
          Resume attached: ${esc(resume.name)}
        </p>
      </div>`;

    const personalization = {
      to: [{ email: TO, name: 'OE Recruitment' }],
      reply_to: { email, name },
    };
    if (env.DKIM_DOMAIN && env.DKIM_SELECTOR && env.DKIM_PRIVATE_KEY) {
      personalization.dkim_domain = env.DKIM_DOMAIN;
      personalization.dkim_selector = env.DKIM_SELECTOR;
      personalization.dkim_private_key = env.DKIM_PRIVATE_KEY;
    }

    const payload = {
      personalizations: [personalization],
      from: { email: FROM, name: 'OE Careers' },
      subject: `Application — ${position} — ${name}`,
      content: [{ type: 'text/html', value: html }],
      attachments: [
        {
          filename: resume.name,
          content: b64,
          type: resume.type || 'application/octet-stream',
          disposition: 'attachment',
        },
      ],
    };

    const send = await fetch('https://api.mailchannels.net/tx/v1/send', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });

    if (!send.ok) {
      const detail = await send.text();
      console.error('MailChannels error', send.status, detail);
      return json({ ok: false, error: 'Mail delivery failed.' }, 502);
    }

    return json({ ok: true });
  } catch (err) {
    console.error('apply handler error', err);
    return json({ ok: false, error: 'Server error.' }, 500);
  }
}

export async function onRequest() {
  return new Response('Method Not Allowed', { status: 405 });
}
