const fs = require('fs');
const path = require('path');
const { checkAdminAuth } = require('../_auth');

module.exports = async function handler(req, res) {
  const clientId = req.query.clientId;
  if (!clientId || !/^[a-z0-9_-]+$/i.test(clientId)) {
    return res.status(400).json({ error: 'Invalid client id' });
  }

  const clientsDir = path.join(process.cwd(), 'clients');
  const filePath = path.join(clientsDir, `${clientId}.json`);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'Client not found' });

  // GET — public, returns portalConfig directly (used by portal.js)
  if (req.method === 'GET') {
    try {
      const client = JSON.parse(fs.readFileSync(filePath, 'utf8'));
      const portalConfig = client.portalConfig || null;
      if (!portalConfig || portalConfig.status === 'off') {
        return res.status(404).json({ error: 'Portal not active' });
      }
      return res.json(portalConfig);
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  // POST / PUT — admin only, saves full portalConfig
  if (req.method === 'POST' || req.method === 'PUT') {
    if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });
    try {
      const client = JSON.parse(fs.readFileSync(filePath, 'utf8'));
      client.portalConfig = req.body;
      fs.writeFileSync(filePath, JSON.stringify(client, null, 2), 'utf8');
      return res.json({ ok: true, client_id: clientId });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  res.status(405).end();
};
