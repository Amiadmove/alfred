const fs = require('fs');
const path = require('path');

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).end();

  const adminPassword = process.env.ADMIN_PASSWORD;
  if (adminPassword) {
    const auth = req.headers['x-admin-password'];
    if (auth !== adminPassword) return res.status(401).json({ error: 'Unauthorized' });
  }

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
    res.json(clients);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
