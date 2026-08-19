const { list } = require('@vercel/blob');

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).end();

  const adminPassword = process.env.ADMIN_PASSWORD;
  if (adminPassword) {
    const auth = req.headers['x-admin-password'];
    if (auth !== adminPassword) return res.status(401).json({ error: 'Unauthorized' });
  }

  try {
    const token = process.env.BLOB_READ_WRITE_TOKEN;
    const { blobs } = await list({ prefix: 'submissions/', limit: 1000, token });
    const jsonBlobs = blobs.filter(b => b.pathname.endsWith('.json'));

    const submissions = await Promise.all(
      jsonBlobs.map(async b => {
        try {
          const resp = await fetch(b.downloadUrl || b.url, {
            headers: { Authorization: `Bearer ${token}` },
          });
          return await resp.json();
        } catch {
          return null;
        }
      })
    );

    const valid = submissions
      .filter(Boolean)
      .sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt));

    res.json(valid);
  } catch (err) {
    console.error(err);
    res.status(500).json([]);
  }
};
