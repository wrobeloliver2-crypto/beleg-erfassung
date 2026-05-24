const https = require('https');

function httpsRequest(method, hostname, path, headers, body) {
  return new Promise((resolve, reject) => {
    const bodyStr = body ? JSON.stringify(body) : '';
    const options = {
      hostname, path, method,
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(bodyStr),
        ...headers
      }
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(data) }); }
        catch(e) { resolve({ status: res.statusCode, body: data }); }
      });
    });
    req.on('error', reject);
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}

async function getAccessToken(tenantId, clientId, clientSecret) {
  const body = `grant_type=client_credentials&client_id=${clientId}&client_secret=${encodeURIComponent(clientSecret)}&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default`;
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.access_token) resolve(parsed.access_token);
          else reject(new Error('Token error: ' + data));
        } catch(e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

exports.handler = async function(event) {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  try {
    const { to, subject, body, pdfBase64, filename } = JSON.parse(event.body);

    const tenantId    = process.env.AZURE_TENANT_ID;
    const clientId    = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    const sender      = process.env.MAIL_SENDER;

    if (!tenantId || !clientId || !clientSecret || !sender) {
      return { statusCode: 500, body: JSON.stringify({ error: 'Missing env vars' }) };
    }

    const token = await getAccessToken(tenantId, clientId, clientSecret);

    const mailPayload = {
      message: {
        subject,
        body: { contentType: 'HTML', content: body },
        toRecipients: [{ emailAddress: { address: to } }],
        attachments: [{
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: filename,
          contentType: 'application/pdf',
          contentBytes: pdfBase64
        }]
      },
      saveToSentItems: true
    };

    const result = await httpsRequest(
      'POST',
      'graph.microsoft.com',
      `/v1.0/users/${sender}/sendMail`,
      { Authorization: `Bearer ${token}` },
      mailPayload
    );

    if (result.status === 202) {
      return { statusCode: 200, body: JSON.stringify({ success: true }) };
    } else {
      return { statusCode: result.status, body: JSON.stringify({ error: result.body }) };
    }

  } catch(err) {
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
