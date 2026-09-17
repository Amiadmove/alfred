module.exports = async function handler(req, res) {
  const pw = process.env.ADMIN_PASSWORD || '';
  res.json({ len: pw.length, first4: pw.substring(0,4) });
};
