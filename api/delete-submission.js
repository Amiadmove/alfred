const { del, list } = require('@vercel/blob');

module.exports = async function handler(req, res) {
  if (req.method !== 'DELETE') return res.status(405).end();

  const { checkAdminAuth } = require('./_auth');
  if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });

  const { id } = req.query;
  if (!id) return res.status(400).json({ error: 'Missing id' });

  try {
    const token = process.env.BLOB_READ_WRITE_TOKEN;
    const { blobs } = await list({ prefix: `submissions/${id}`, limit: 10, token });
    if (blobs.length > 0) {
      await del(blobs.map(b => b.url), { token });
    }
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
};
