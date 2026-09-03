const { list, put, del } = require('@vercel/blob');

module.exports = async function handler(req, res) {
  const { checkAdminAuth } = require('./_auth');
  if (!checkAdminAuth(req)) return res.status(401).json({ error: 'Unauthorized' });

  const token = process.env.BLOB_READ_WRITE_TOKEN;

  // GET /api/documents?id=... → list documents
  // GET /api/documents?url=... → download document
  if (req.method === 'GET') {
    const { id, url } = req.query;

    if (url) {
      try {
        const upstream = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
        if (!upstream.ok) return res.status(404).end();
        const ct = upstream.headers.get('content-type') || 'application/octet-stream';
        res.setHeader('Content-Type', ct);
        res.send(Buffer.from(await upstream.arrayBuffer()));
      } catch (err) {
        console.error(err);
        res.status(500).end();
      }
      return;
    }

    if (id) {
      try {
        const { blobs } = await list({ prefix: `submissions/${id}/docs/`, token });
        res.json(blobs.map(b => ({
          name: b.pathname.split('/').pop(),
          url: b.url,
          size: b.size,
          uploadedAt: b.uploadedAt,
        })));
      } catch (err) {
        console.error(err);
        res.status(500).json({ ok: false, error: err.message });
      }
      return;
    }

    return res.status(400).json({ error: 'Missing id or url' });
  }

  // POST /api/documents?id=&filename= → upload document
  if (req.method === 'POST') {
    const { id, filename } = req.query;
    if (!id || !filename) return res.status(400).json({ error: 'Missing id or filename' });
    try {
      const { content, contentType } = req.body;
      const buffer = Buffer.from(content, 'base64');
      const blob = await put(`submissions/${id}/docs/${filename}`, buffer, {
        access: 'private',
        contentType: contentType || 'application/octet-stream',
        token,
        allowOverwrite: true,
      });
      res.json({ ok: true, url: blob.url });
    } catch (err) {
      console.error(err);
      res.status(500).json({ ok: false, error: err.message });
    }
    return;
  }

  // DELETE /api/documents?url= → delete document
  if (req.method === 'DELETE') {
    const { url } = req.query;
    if (!url) return res.status(400).json({ error: 'Missing url' });
    try {
      await del(url, { token });
      res.json({ ok: true });
    } catch (err) {
      console.error(err);
      res.status(500).json({ ok: false, error: err.message });
    }
    return;
  }

  res.status(405).end();
};
