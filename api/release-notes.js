const { put, list } = require('@vercel/blob');
const { checkAdminAuth } = require('./_auth');

const BLOB_PATH = 'release-notes/global.json';

async function readFromBlob() {
  if (!process.env.BLOB_READ_WRITE_TOKEN) return null;
  try {
    const { blobs } = await list({ prefix: BLOB_PATH, token: process.env.BLOB_READ_WRITE_TOKEN });
    const blob = blobs.find(b => b.pathname === BLOB_PATH);
    if (!blob) return null;
    const r = await fetch(blob.downloadUrl || blob.url);
    if (!r.ok) return null;
    return await r.json();
  } catch { return null; }
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  // GET — public
  if (req.method === 'GET') {
    try {
      const data = await readFromBlob();
      return res.json(data || { items: [] });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  // PUT — admin only
  if (req.method === 'PUT') {
    if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });
    if (!process.env.BLOB_READ_WRITE_TOKEN) return res.status(503).json({ error: 'Blob not configured' });
    try {
      await put(BLOB_PATH, JSON.stringify(req.body, null, 2), {
        access: 'private',
        addRandomSuffix: false,
        contentType: 'application/json',
        token: process.env.BLOB_READ_WRITE_TOKEN,
      });
      return res.json({ ok: true });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  res.status(405).end();
};
