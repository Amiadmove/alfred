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

  // GET — public, returns portal config if active
  if (req.method === 'GET') {
    try {
      const client = JSON.parse(fs.readFileSync(filePath, 'utf8'));
      const portalConfig = client.portalConfig || null;
      if (!portalConfig || !portalConfig.active) {
        return res.status(404).json({ error: 'Portal not active' });
      }
      return res.json({ id: client.id, name: client.name, portalConfig });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  // POST — admin only, saves portal config
  if (req.method === 'POST') {
    if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });
    try {
      const client = JSON.parse(fs.readFileSync(filePath, 'utf8'));
      client.portalConfig = req.body;
      fs.writeFileSync(filePath, JSON.stringify(client, null, 2), 'utf8');
      return res.json({ ok: true });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  res.status(405).end();
};
