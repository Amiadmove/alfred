const fs = require('fs');
const path = require('path');
const { checkAdminAuth } = require('../_auth');
const { put, list } = require('@vercel/blob');

const BLOB_PREFIX = 'portal-clients/';

// Read client JSON from Vercel Blob (returns null if not found or Blob not configured)
async function readFromBlob(clientId) {
  if (!process.env.BLOB_READ_WRITE_TOKEN) return null;
  try {
    const { blobs } = await list({
      prefix: BLOB_PREFIX + clientId + '.json',
      token: process.env.BLOB_READ_WRITE_TOKEN,
    });
    const blob = blobs.find(b => b.pathname === BLOB_PREFIX + clientId + '.json');
    if (!blob) return null;
    // For private stores use downloadUrl (signed), fall back to url
    const fetchUrl = blob.downloadUrl || blob.url;
    const r = await fetch(fetchUrl);
    if (!r.ok) return null;
    return await r.json();
  } catch {
    return null;
  }
}

// Read client JSON from local filesystem (git source of truth for initial data)
function readFromFile(clientId) {
  const filePath = path.join(process.cwd(), 'clients', `${clientId}.json`);
  if (!fs.existsSync(filePath)) return null;
  try { return JSON.parse(fs.readFileSync(filePath, 'utf8')); } catch { return null; }
}

module.exports = async function handler(req, res) {
  const clientId = req.query.clientId;
  if (!clientId || !/^[a-z0-9_-]+$/i.test(clientId)) {
    return res.status(400).json({ error: 'Invalid client id' });
  }

  // GET — public, returns portalConfig (used by portal.js)
  if (req.method === 'GET') {
    try {
      // Blob takes priority (latest saved data), fall back to git file
      const client = (await readFromBlob(clientId)) || readFromFile(clientId);
      if (!client) return res.status(404).json({ error: 'Client not found' });
      const portalConfig = client.portalConfig || null;
      if (!portalConfig || portalConfig.status === 'off') {
        return res.status(404).json({ error: 'Portal not active' });
      }
      return res.json(portalConfig);
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  // POST / PUT — admin only, saves portalConfig to Vercel Blob
  if (req.method === 'POST' || req.method === 'PUT') {
    if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });
    if (!process.env.BLOB_READ_WRITE_TOKEN) {
      return res.status(503).json({ error: 'Vercel Blob not configured. Add BLOB_READ_WRITE_TOKEN to environment variables.' });
    }
    try {
      // Start with git file as base, overlay with any previously saved Blob data
      const base = readFromFile(clientId);
      if (!base) return res.status(404).json({ error: 'Client not found' });
      const saved = await readFromBlob(clientId);
      const client = saved || base;
      client.portalConfig = req.body;

      await put(BLOB_PREFIX + clientId + '.json', JSON.stringify(client, null, 2), {
        access: 'private',
        addRandomSuffix: false,
        allowOverwrite: true,
        contentType: 'application/json',
        token: process.env.BLOB_READ_WRITE_TOKEN,
      });

      return res.json({ ok: true, client_id: clientId });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  res.status(405).end();
};
