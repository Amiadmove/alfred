const { put, del, list } = require('@vercel/blob');
const { checkAdminAuth } = require('./_auth');
const crypto = require('crypto');

const PREFIX = 'alfred-templates/';

module.exports = async function handler(req, res) {
  const { action, id } = req.query;

  // ── GET: list all templates ──────────────────────────────────────────────
  if (req.method === 'GET' && action === 'list') {
    try {
      const { blobs } = await list({ prefix: PREFIX });
      const templates = await Promise.all(
        blobs.map(async (blob) => {
          const r = await fetch(blob.url);
          return r.json();
        })
      );
      templates.sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
      return res.json({ ok: true, templates });
    } catch (e) {
      return res.status(500).json({ ok: false, error: e.message });
    }
  }

  // ── GET: fetch single template by id ────────────────────────────────────
  if (req.method === 'GET' && action === 'get' && id) {
    try {
      const { blobs } = await list({ prefix: PREFIX + id + '.' });
      if (!blobs.length) return res.status(404).json({ ok: false, error: 'Template not found' });
      const r = await fetch(blobs[0].url);
      const data = await r.json();
      return res.json({ ok: true, data });
    } catch (e) {
      return res.status(500).json({ ok: false, error: e.message });
    }
  }

  // ── POST: save / update template (admin only) ────────────────────────────
  if (req.method === 'POST') {
    if (!checkAdminAuth(req)) return res.status(401).json({ ok: false, error: 'Unauthorized' });

    const body = req.body;
    if (!body || !body.name || !body.formType) {
      return res.status(400).json({ ok: false, error: 'Missing name or formType' });
    }

    try {
      const templateId = body.id || crypto.randomBytes(6).toString('hex');
      const template = {
        id: templateId,
        name: body.name,
        formType: body.formType,   // 'holidayheroes' | 'ratecore'
        sections: body.sections || {},
        createdAt: body.createdAt || new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };

      await put(PREFIX + templateId + '.json', JSON.stringify(template), {
        access: 'public',
        contentType: 'application/json',
        addRandomSuffix: false,
        allowOverwrite: true,
      });

      return res.json({ ok: true, id: templateId, template });
    } catch (e) {
      return res.status(500).json({ ok: false, error: e.message });
    }
  }

  // ── DELETE: remove template (admin only) ────────────────────────────────
  if (req.method === 'DELETE') {
    if (!checkAdminAuth(req)) return res.status(401).json({ ok: false, error: 'Unauthorized' });
    if (!id) return res.status(400).json({ ok: false, error: 'Missing id' });

    try {
      const { blobs } = await list({ prefix: PREFIX + id + '.' });
      if (!blobs.length) return res.status(404).json({ ok: false, error: 'Template not found' });
      await del(blobs[0].url);
      return res.json({ ok: true });
    } catch (e) {
      return res.status(500).json({ ok: false, error: e.message });
    }
  }

  return res.status(405).json({ ok: false, error: 'Method not allowed' });
};
