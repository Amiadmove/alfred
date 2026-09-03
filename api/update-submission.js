const { put } = require('@vercel/blob');

module.exports = async function handler(req, res) {
  if (req.method !== 'PUT') return res.status(405).end();

  const { checkAdminAuth } = require('./_auth');
  if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });

  const { id } = req.query;
  if (!id) return res.status(400).json({ error: 'Missing id' });

  try {
    const token = process.env.BLOB_READ_WRITE_TOKEN;
    await put(`submissions/${id}.json`, JSON.stringify(req.body), {
      access: 'private',
      contentType: 'application/json',
      token,
      allowOverwrite: true,
    });
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
};
