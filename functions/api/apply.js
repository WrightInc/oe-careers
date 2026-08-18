/**
 * POST /api/apply — Hydrovac Operator application handler
 *
 * Cloudflare Pages Function. Accepts multipart/form-data, validates,
 * and emails the application to OE with the resume attached, sent through
 * OE's own Microsoft 365 tenant via the Microsoft Graph API.
 *
 * Required environment variables (Cloudflare Pages → Settings → Environment variables):
 *   TENANT_ID      Directory (tenant) ID from Entra
 *   CLIENT_ID      Application (client) ID from Entra
 *   CLIENT_SECRET  Client secret VALUE (not the Secret ID) — store as encrypted
 *   SEND_MAILBOX   Mailbox the app is authorised to send as, e.g. opportunities@oeservices.ca
 *   TO_EMAIL       Where applications are delivered, e.g. opportunities@oeservices.ca
 *
 * The app registration must have Mail.Send granted through Exchange RBAC and
 * scoped to SEND_MAILBOX only. It must NOT have Mail.Send consented in Entra,
 * as the two grants combine and would remove the mailbox restriction.
 */

// Graph caps the whole JSON request at 4 MB. Base64 inflates a file by up to
// 33%, so the raw file limit is held well below that.
const MAX_BYTES = 2 * 1024 * 1024;
const MAX_LABEL = '2 MB';
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

async function getGraphToken(env) {
  const body = new URLSearchParams({
    client_id: env.CLIENT_ID,
    client_secret: env.CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body,
    }
  );

  if (!res.ok) {
    const detail = await res.text();
    console.error('Graph token error', res.status, detail);
    return null;
  }

  const data = await res.json();
  return data.access_token || null;
}

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
    const hasResume =
      resume && typeof resume.arrayBuffer === 'function' && resume.size > 0;

    let b64 = null;
    if (hasResume) {
      const ext = (resume.name || '').split('.').pop().toLowerCase();
      if (!ALLOWED_EXT.includes(ext)) {
        return json({ ok: false, error: 'Unsupported file type.' }, 400);
      }

      const bytes = await resume.arrayBuffer();
      if (bytes.byteLength > MAX_BYTES) {
        return json({ ok: false, error: `File exceeds ${MAX_LABEL}.` }, 400);
      }

      // base64 encode in chunks (avoids call-stack overflow on larger files)
      const view = new Uint8Array(bytes);
      let binary = '';
      for (let i = 0; i < view.length; i += 0x8000) {
        binary += String.fromCharCode.apply(null, view.subarray(i, i + 0x8000));
      }
      b64 = btoa(binary);
    }

    const TO = env.TO_EMAIL || 'opportunities@oeservices.ca';
    const SENDER = env.SEND_MAILBOX || TO;

    if (!env.TENANT_ID || !env.CLIENT_ID || !env.CLIENT_SECRET) {
      console.error('Graph credentials not configured');
      return json({ ok: false, error: 'Mail delivery failed.' }, 502);
    }

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
          ${hasResume ? `Resume attached: ${esc(resume.name)}` : 'No resume attached.'}
        </p>
      </div>`;

    const token = await getGraphToken(env);
    if (!token) {
      return json({ ok: false, error: 'Mail delivery failed.' }, 502);
    }

    const payload = {
      message: {
        subject: `Application — ${position} — ${name}`,
        body: { contentType: 'HTML', content: html },
        toRecipients: [{ emailAddress: { address: TO } }],
        replyTo: [{ emailAddress: { address: email, name } }],
        attachments: hasResume
          ? [
              {
                '@odata.type': '#microsoft.graph.fileAttachment',
                name: resume.name,
                contentType: resume.type || 'application/octet-stream',
                contentBytes: b64,
              },
            ]
          : [],
      },
      saveToSentItems: false,
    };

    const send = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(SENDER)}/sendMail`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload),
      }
    );

    if (!send.ok) {
      const detail = await send.text();
      console.error('Graph sendMail error', send.status, detail);
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
