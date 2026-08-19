module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).end();

  const adminPassword = process.env.ADMIN_PASSWORD;
  if (adminPassword) {
    const auth = req.headers['x-admin-password'];
    if (auth !== adminPassword) return res.status(401).end();
  }

  const { url } = req.query;
  if (!url) return res.status(400).json({ error: 'Missing url parameter' });

  try {
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${process.env.BLOB_READ_WRITE_TOKEN}` },
    });

    if (!resp.ok) return res.status(resp.status).end();

    const buffer = await resp.arrayBuffer();
    const filename = url.split('/').pop().split('?')[0] || 'submission.xlsx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error(err);
    res.status(500).end();
  }
};
