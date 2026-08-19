const { put } = require('@vercel/blob');

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const adminPassword = process.env.ADMIN_PASSWORD;
  if (adminPassword && req.headers['x-admin-password'] !== adminPassword) return res.status(401).end();

  const { id, filename } = req.query;
  if (!id || !filename) return res.status(400).json({ error: 'Missing id or filename' });

  try {
    const { content, contentType } = req.body;
    const buffer = Buffer.from(content, 'base64');
    const token = process.env.BLOB_READ_WRITE_TOKEN;
    const blob = await put(`submissions/${id}/docs/${filename}`, buffer, {
      access: 'private',
      contentType: contentType || 'application/octet-stream',
      token,
      allowOverwrite: true,
    });
    res.json({ ok: true, url: blob.url });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
};
