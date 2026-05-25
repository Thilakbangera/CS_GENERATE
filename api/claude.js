// Vercel serverless function — proxies Claude API requests to avoid CORS
// Equivalent to the Vite dev-server proxy at /api/claude → https://api.anthropic.com
export default async function handler(req, res) {
  // Only allow POST
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const anthropicRes = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type':      'application/json',
        'anthropic-version': req.headers['anthropic-version'] || '2023-06-01',
        'x-api-key':         req.headers['x-api-key'] || '',
      },
      body: JSON.stringify(req.body),
    });

    const data = await anthropicRes.json();
    return res.status(anthropicRes.status).json(data);
  } catch (err) {
    return res.status(500).json({ error: err.message || 'Proxy error' });
  }
}
