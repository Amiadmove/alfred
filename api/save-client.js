const fs = require('fs');
const path = require('path');

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { checkAdminAuth } = require('./_auth');
  if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });

  const data = req.body;
  if (!data || !data.id || !/^[a-z0-9_-]+$/i.test(data.id)) {
    return res.status(400).json({ error: 'Invalid client id' });
  }

  const clientsDir = path.join(process.cwd(), 'clients');
  if (!fs.existsSync(clientsDir)) fs.mkdirSync(clientsDir, { recursive: true });
  const filePath = path.join(clientsDir, `${data.id}.json`);

  try {
    fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
