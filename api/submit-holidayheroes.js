const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const { put } = require('@vercel/blob');

const NOTIFY_EMAIL = process.env.NOTIFY_EMAIL || 'amiad@alfredtravel.io';

function buildExcel(d) {
  const wb = xlsx.utils.book_new();

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['MOVE — HOLIDAYHEROES PARTNER ONBOARDING QUESTIONNAIRE'],
    ['HolidayHeroes Platform  ·  Powered by Move'],
    [],
    ['Completed on', new Date().toLocaleDateString('en-GB')],
  ]), 'Instructions');

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['COMPANY & CONTACT INFORMATION'],
    [],
    ['Company Name',  d.company?.name    || ''],
    ['Website',       d.company?.website  || ''],
    ['Country / HQ',  d.company?.country  || ''],
    ['Company Type',  d.company?.type     || ''],
    [],
    ['COMMERCIAL CONTACT'],
    ['Full Name',  d.commercial?.name  || ''],
    ['Job Title',  d.commercial?.title || ''],
    ['Email',      d.commercial?.email || ''],
    ['Phone',      d.commercial?.phone || ''],
    [],
    ['TECHNICAL CONTACT'],
    ['Full Name',  d.technical?.name  || ''],
    ['Job Title',  d.technical?.title || ''],
    ['Email',      d.technical?.email || ''],
    ['Phone',      d.technical?.phone || ''],
  ]), '1 — Company & Contacts');

  const routeRows = (d.routes || []).filter(r => r.name)
    .map((r, i) => [i + 1, r.name, r.domestic ? 'Yes' : '', r.international ? 'Yes' : '']);

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['MARKET & ROUTES'],
    [],
    ['Q1 — Operating Market',             d.market             || ''],
    ['Q2 — Selling Currencies',           d.currencies         || ''],
    ['Q3 — Departure Airports',           d.departureAirports  || ''],
    [],
    ['Q4 — Key Destinations (80% of revenue)'],
    [d.destinations || ''],
    [],
    ['Q5 — Top 20 Routes'],
    ['#', 'Route / Destination', 'Domestic', 'International'],
    ...routeRows,
    [],
    ['Q6 — Total Routes Currently Selling', d.totalRoutes    || ''],
    ['Q7 — % Routes Making 80% Revenue',    d.revenueRoutePct || ''],
  ]), '2 — Market & Routes');

  const airRows   = (d.airContent   || []).filter(a => a.airline).map(a => [a.airline, a.gds, a.tf, a.aggregator, a.direct, a.ndc, a.negotiated]);
  const hotelRows = (d.hotelContent || []).filter(h => h.name).map(h   => [h.name, h.gds, h.bedbank, h.directConnect, h.aggregator, h.other]);
  const nonAirRows= (d.nonAirContent|| []).filter(n => n.type).map(n   => [n.type, n.source, n.negotiated]);

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['CONTENT SOURCES'],
    [],
    ['Q8 — Air Content'],
    ['Airline','GDS','TF','Aggregator','Direct','NDC','Negotiated Pricing'],
    ...airRows,
    [],
    ['Q9 — Hotel Content'],
    ['Name / Group','GDS','Bedbank(s)','Direct Connect','Aggregator / Channel Mgr','Other'],
    ...hotelRows,
    [],
    ['Q10 — Non-Air Content'],
    ['Type','Source / Connection','Negotiated Pricing'],
    ...nonAirRows,
  ]), '3 — Content');

  const tourOpRows = (d.tourOperators || []).filter(t => t.name).map(t => [t.name, t.connection]);

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['OPERATIONS & TECHNOLOGY'],
    [],
    ['Q11 — Resell Tour Operators?', d.resellTourOps || ''],
    ...(tourOpRows.length ? [['Tour Operator','Connection'], ...tourOpRows, []] : [[]]),
    ['Q12 — Payment Gateway',        d.paymentGateway  || ''],
    ['Q13 — Paying Party',           d.payingParty     || ''],
    ['Q14 — Payment Terms',          d.paymentTerms    || ''],
    ['Q15 — Back Office',            d.backOffice      || ''],
    [],
    ['Q16 — Current Applications'],
    ['B2B', d.currentApps?.b2b||''],['B2C', d.currentApps?.b2c||''],['B2B2C', d.currentApps?.b2b2c||''],
    [],
    ['Q17 — Future Applications'],
    ['B2B', d.futureApps?.b2b||''],['B2C', d.futureApps?.b2c||''],['B2B2C', d.futureApps?.b2b2c||''],
    [],
    ['Q18 — Replacing Applications'],
    ['B2B', d.replacingApps?.b2b||''],['B2C', d.replacingApps?.b2c||''],['B2B2C', d.replacingApps?.b2b2c||''],
    [],
    ['Q19 — B2B2C Partners',         d.b2b2cPartners   || ''],
    ['Q20 — Merchant of Record',     d.merchantOfRecord|| ''],
    [],
    ['Q21 — B2B Agents'],
    ['Nr of Sub Agents',             d.b2bAgents?.count || ''],
    ['% Producing 50+ bookings/yr',  d.b2bAgents?.pct   || ''],
    [],
    ['Q22 — UI/UX Requirements',     d.uiRequirements  || ''],
    ['Q23 — Other Technical Info',   d.otherInfo       || ''],
  ]), '4 — Operations & Tech');

  const hotelSupplierRows = (d.providers?.hotelSuppliers || []).filter(s => s.name)
    .map(s => [s.name, s.apiEndpoint, s.username, '(redacted)', s.notes]);

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['PROVIDER CREDENTIALS'],
    [],
    ['AMADEUS'],
    ['OID Client ID',     d.providers?.amadeus?.oidClientId    || ''],
    ['OID Client Secret', '(redacted)'],
    ['Environment',       d.providers?.amadeus?.oidEnvironment || ''],
    ['Notes',             d.providers?.amadeus?.notes          || ''],
    [],
    ['TRAVELFUSION'],
    ['IP Whitelist Requested', d.providers?.travelfusion?.ipWhitelistRequested || ''],
    ['TF Username',            d.providers?.travelfusion?.tfUsername || ''],
    ['TF Password',            '(redacted)'],
    ['Notes',                  d.providers?.travelfusion?.notes || ''],
    [],
    ['HOTEL SUPPLIERS'],
    ['Gimmonix', d.providers?.gimmonix || ''],
    [],
    ...(hotelSupplierRows.length ? [
      ['Supplier Name', 'API Endpoint', 'Username', 'Password', 'Notes'],
      ...hotelSupplierRows,
    ] : [['No additional hotel suppliers listed']]),
  ]), '5 — Provider Credentials');

  return xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  try {
    const d = req.body;
    const partnerName = d.company?.name || 'Unknown';
    const id = Date.now().toString();
    const buffer = buildExcel(d);
    const filename = `HH_${partnerName.replace(/[^a-z0-9]/gi, '_')}_${new Date().toISOString().slice(0,10)}.xlsx`;

    const submissionData = {
      id,
      platform: 'holidayheroes',
      partnerName,
      submittedAt: new Date().toISOString(),
      data: d,
    };

    // Save to Vercel Blob (optional — requires BLOB_READ_WRITE_TOKEN)
    if (process.env.BLOB_READ_WRITE_TOKEN) {
      try {
        const xlsxBlob = await put(`submissions/${id}.xlsx`, Buffer.from(buffer), {
          access: 'public',
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        submissionData.excelUrl = xlsxBlob.url;
        await put(`submissions/${id}.json`, JSON.stringify(submissionData), {
          access: 'public',
          contentType: 'application/json',
        });
      } catch (blobErr) {
        console.error('Blob storage error (non-fatal):', blobErr.message);
      }
    }

    // Send email
    if (process.env.SMTP_USER && process.env.SMTP_PASS) {
      const transporter = nodemailer.createTransport({
        host: process.env.SMTP_HOST || 'smtp.gmail.com',
        port: Number(process.env.SMTP_PORT) || 587,
        secure: false,
        auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
      });
      await transporter.sendMail({
        from: `"Move Onboarding" <${process.env.SMTP_USER}>`,
        to: NOTIFY_EMAIL,
        subject: `New HolidayHeroes Onboarding — ${partnerName}`,
        text: `New HolidayHeroes partner onboarding submitted.\n\nPartner: ${partnerName}\nMarket: ${d.market || '—'}\nContact: ${d.commercial?.name || '—'} (${d.commercial?.email || '—'})\n\nExcel file attached.`,
        attachments: [{ filename, content: buffer }],
      });
    }

    res.status(200).json({ ok: true, filename });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
};
