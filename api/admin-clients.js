const fs = require('fs');
const path = require('path');
const { checkAdminAuth } = require('./_auth');

module.exports = async function handler(req, res) {
  if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });

  // GET — list all clients
  if (req.method === 'GET') {
    try {
      const clientsDir = path.join(process.cwd(), 'clients');
      const files = fs.readdirSync(clientsDir).filter(f => f.endsWith('.json'));
      const clients = files.map(f => {
        try {
          return JSON.parse(fs.readFileSync(path.join(clientsDir, f), 'utf8'));
        } catch {
          return null;
        }
      }).filter(Boolean);
      return res.json(clients);
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  // POST — save a client
  if (req.method === 'POST') {
    const data = req.body;
    if (!data || !data.id || !/^[a-z0-9_-]+$/i.test(data.id)) {
      return res.status(400).json({ error: 'Invalid client id' });
    }
    const clientsDir = path.join(process.cwd(), 'clients');
    if (!fs.existsSync(clientsDir)) fs.mkdirSync(clientsDir, { recursive: true });
    const filePath = path.join(clientsDir, `${data.id}.json`);
    try {
      fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
      return res.json({ ok: true });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  res.status(405).end();
};
