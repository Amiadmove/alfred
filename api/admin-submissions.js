const { list } = require('@vercel/blob');

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).end();

  try {
    const { blobs } = await list({ prefix: 'submissions/', limit: 1000 });
    const jsonBlobs = blobs.filter(b => b.pathname.endsWith('.json'));

    const submissions = await Promise.all(
      jsonBlobs.map(async b => {
        try {
          const resp = await fetch(b.url, {
            headers: { Authorization: `Bearer ${process.env.BLOB_READ_WRITE_TOKEN}` },
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
