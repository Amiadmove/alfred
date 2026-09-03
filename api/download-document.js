module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).end();

  const { checkAdminAuth } = require('./_auth');
  if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });

  const { url } = req.query;
  if (!url) return res.status(400).json({ error: 'Missing url' });

  try {
    const token = process.env.BLOB_READ_WRITE_TOKEN;
    const upstream = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!upstream.ok) return res.status(404).end();
    const ct = upstream.headers.get('content-type') || 'application/octet-stream';
    res.setHeader('Content-Type', ct);
    const buf = await upstream.arrayBuffer();
    res.send(Buffer.from(buf));
  } catch (err) {
    console.error(err);
    res.status(500).end();
  }
};
