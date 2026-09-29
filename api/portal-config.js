const { put, list } = require('@vercel/blob');
const { checkAdminAuth } = require('./_auth');
const fs = require('fs');
const path = require('path');

const TOKEN = process.env.BLOB_READ_WRITE_TOKEN;

function blobPath(id) {
  return `portal-configs/${id}.json`;
}

async function readFromBlob(id) {
  if (!TOKEN) return null;
  try {
    const bp = blobPath(id);
    const { blobs } = await list({ prefix: bp, token: TOKEN });
    const blob = blobs.find(b => b.pathname === bp);
    if (!blob) return null;
    const r = await fetch(blob.url, { headers: { Authorization: `Bearer ${TOKEN}` } });
    if (!r.ok) return null;
    return await r.json();
  } catch { return null; }
}

function readFromFile(id) {
  try {
    const filePath = path.join(process.cwd(), 'clients', `${id}.json`);
    const raw = fs.readFileSync(filePath, 'utf8');
    const client = JSON.parse(raw);
    return client.portalConfig || null;
  } catch { return null; }
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, PUT, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-admin-password, x-session-token');

  if (req.method === 'OPTIONS') return res.status(200).end();

  // Support both /api/portal-config?client=setur and /api/portal-config/setur
  const urlParts = (req.url || '').split('?')[0].split('/').filter(Boolean);
  const lastSegment = urlParts[urlParts.length - 1];
  const id = req.query.client || req.query.id ||
    (lastSegment !== 'portal-config' ? lastSegment : null);

  if (!id || !/^[a-z0-9_-]+$/i.test(id)) {
    return res.status(400).json({ error: 'Missing or invalid client id' });
  }

  // GET — public (portal reads this)
  if (req.method === 'GET') {
    const blobData = await readFromBlob(id);
    if (blobData) return res.json(blobData);
    const fileData = readFromFile(id);
    if (fileData) return res.json(fileData);
    return res.status(404).json({ error: 'Portal config not found' });
  }

  // PUT — admin only
  if (req.method === 'PUT') {
    if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });
    if (!TOKEN) return res.status(503).json({ error: 'Blob not configured' });

    const body = req.body;
    if (!body) return res.status(400).json({ error: 'Missing body' });

    try {
      await put(blobPath(id), JSON.stringify(body, null, 2), {
        access: 'private',
        contentType: 'application/json',
        addRandomSuffix: false,
        allowOverwrite: true,
        token: TOKEN,
      });
      return res.json({ ok: true });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  res.status(405).end();
};
