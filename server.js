const express = require('express');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

// ─── Config ───────────────────────────────────────────────────────────────────
const NOTIFY_EMAIL    = 'amiad@alfredtravel.io';
const SMTP_HOST       = process.env.SMTP_HOST  || 'smtp.gmail.com';
const SMTP_PORT       = process.env.SMTP_PORT  || 587;
const SMTP_USER       = process.env.SMTP_USER;
const SMTP_PASS       = process.env.SMTP_PASS;
const PORT            = process.env.PORT       || 3000;
const SUBMISSIONS_DIR = path.join(__dirname, 'submissions');

if (!fs.existsSync(SUBMISSIONS_DIR)) fs.mkdirSync(SUBMISSIONS_DIR);

// ─── Shared: send email ────────────────────────────────────────────────────────
async function sendEmail(partnerName, platform, buffer, filename) {
  if (!SMTP_USER || !SMTP_PASS) return;
  const transporter = nodemailer.createTransport({
    host: SMTP_HOST, port: SMTP_PORT, secure: false,
    auth: { user: SMTP_USER, pass: SMTP_PASS },
  });
  await transporter.sendMail({
    from: `"Move Onboarding" <${SMTP_USER}>`,
    to: NOTIFY_EMAIL,
    subject: `New Onboarding — ${partnerName} [${platform}]`,
    text: `A new partner onboarding form has been submitted.\n\nPartner: ${partnerName}\nPlatform: ${platform}\n\nThe completed Excel file is attached.`,
    attachments: [{ filename, content: buffer }],
  });
  console.log(`Email sent to ${NOTIFY_EMAIL}`);
}

// ─── Save submission to disk (Excel + JSON) ────────────────────────────────────
function saveSubmission(platform, partnerName, data, excelBuffer) {
  const slug      = (partnerName || 'Unknown').replace(/[^a-z0-9]/gi, '_');
  const timestamp = new Date().toISOString().slice(0, 10);
  const base      = `${platform}_${slug}_${timestamp}`;

  const xlsxPath = path.join(SUBMISSIONS_DIR, `${base}.xlsx`);
  const jsonPath = path.join(SUBMISSIONS_DIR, `${base}.json`);

  fs.writeFileSync(xlsxPath, excelBuffer);
  fs.writeFileSync(jsonPath, JSON.stringify({
    id: base,
    platform,
    partnerName,
    submittedAt: new Date().toISOString(),
    data,
  }, null, 2));

  console.log(`Saved: ${xlsxPath}`);
  return { xlsxPath, jsonPath, filename: `${base}.xlsx` };
}

// ─── Build Excel: Ratecore ────────────────────────────────────────────────────
function buildRatecoreExcel(d) {
  const wb = xlsx.utils.book_new();

  const s0 = [
    ['MOVE — PARTNER ONBOARDING QUESTIONNAIRE'],
    ['Ratecore Integration  ·  Powered by Move'],
    [],
    ['Completed via Move Partner Onboarding portal on', new Date().toLocaleDateString('en-GB')],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s0), 'Instructions');

  const s1 = [
    ['COMPANY & CONTACT INFORMATION'],
    [],
    ['  1.1  Company Information'],
    ['Company Name',               '', d.company.name],
    ['Website',                    '', d.company.website],
    ['Registered Address',         '', d.company.address],
    ['Company Reg. Number',        '', d.company.regNumber],
    ['VAT Number',                 '', d.company.vatNumber],
    ['Primary Market(s)',          '', d.company.primaryMarkets],
    ['App Platform(s)',            '', d.company.appPlatforms],
    [],
    ['  1.2  Commercial Contact'],
    ['Full Name',  '', d.commercial.name],
    ['Job Title',  '', d.commercial.title],
    ['Email',      '', d.commercial.email],
    ['Phone',      '', d.commercial.phone],
    [],
    ['  1.3  Technical Contact'],
    ['Full Name',             '', d.technical.name],
    ['Job Title',             '', d.technical.title],
    ['Email',                 '', d.technical.email],
    ['Phone',                 '', d.technical.phone],
    ['GitHub / GitLab handle','', d.technical.github],
    [],
    ['  1.4  Finance / Billing Contact'],
    ['Full Name',                  '', d.finance.name],
    ['Email',                      '', d.finance.email],
    ['Preferred Invoice Currency', '', d.finance.currency],
    ['PO Number Required?',        '', d.finance.poRequired + (d.finance.poNumber ? ' — ' + d.finance.poNumber : '')],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s1), '1 — Company & Contacts');

  const monthNames = Array.from({ length: 12 }, (_, i) => `Month ${i + 1}`);
  const totalGBV = d.months.reduce((sum, m) => sum + ((Number(m.bookings) || 0) * (Number(m.avgValue) || 0)), 0);
  const sumField = (f) => d.months.reduce((s, m) => s + (Number(m[f]) || 0), 0);

  const s2 = [
    ['TRAFFIC & BOOKING VOLUME FORECAST — 12 MONTHS POST GO-LIVE'],
    [],
    ['Month', 'Monthly App Sessions', 'Travel Searches', 'Confirmed Bookings', 'Avg. Booking Value (EUR)', 'Expected GBV (EUR)', 'Cancellation Rate (%)', 'Notes'],
    ...d.months.map((m, i) => [
      monthNames[i],
      m.sessions || '', m.searches || '', m.bookings || '', m.avgValue || '',
      (m.bookings && m.avgValue) ? Number(m.bookings) * Number(m.avgValue) : '',
      m.cancellationRate || '', m.notes || '',
    ]),
    ['TOTAL (12 mo.)', sumField('sessions'), sumField('searches'), sumField('bookings'), '', totalGBV, '', ''],
    [],
    ['  Additional Context'],
    ['Peak Travel Months',         '', d.forecast.peakMonths],
    ['Seasonal Patterns',          '', d.forecast.seasonalPatterns],
    ['Avg. Lead Time to Departure','', d.forecast.avgLeadTime],
    ['Average Trip Duration',      '', d.forecast.avgTripDuration],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s2), '2 — Volume Forecast');

  const airportTotal = d.airports.reduce((s, a) => s + (Number(a.traffic) || 0), 0);
  const destTotal    = d.destinations.reduce((s, dest) => s + (Number(dest.traffic) || 0), 0);
  const s3 = [
    ['ROUTES, DESTINATIONS & USER SEGMENTS'],
    [],
    ['  3.1  Top Departure Airports'],
    ['#', 'Airport Name', 'IATA Code', 'City / Country', 'Est. % of Traffic', 'Notes'],
    ...d.airports.map((a, i) => [i + 1, a.name, a.iata, a.cityCountry, a.traffic || '', a.notes]),
    ['Total', '', '', '', airportTotal + '%', ''],
    [],
    ['  3.2  Top Destinations'],
    ['#', 'City / Region', 'Country', 'Product Type', 'Est. % of Traffic', 'Notes'],
    ...d.destinations.map((dest, i) => [i + 1, dest.cityRegion, dest.country, dest.productType, dest.traffic || '', dest.notes]),
    ['Total', '', '', '', destTotal + '%', ''],
    [],
    ['  3.3  User Segments & Travel Profile'],
    ['Primary User Segments',    '', '', d.segments.primarySegments],
    ['Typical Group Size',       '', '', d.segments.groupSize],
    ['Price Sensitivity',        '', '', d.segments.priceSensitivity],
    ['Booking Preference',       '', '', d.segments.bookingPreference],
    ['Avg. Days to Departure',   '', '', d.segments.avgDaysToDeparture],
    ['Mobile vs Desktop Split',  '', '', d.segments.mobileDesktopSplit],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s3), '3 — Routes & Markets');

  const s4 = [
    ['TECHNICAL INTEGRATION DETAILS'],
    [],
    ['  4.1  Tech Stack'],
    ['Frontend Framework',       '', '', d.tech.frontendFramework],
    ['Backend Language',         '', '', d.tech.backendLanguage],
    ['API Preference',           '', '', d.tech.apiPreference],
    ['Authentication Method',    '', '', d.tech.authMethod],
    ['Hosting / Cloud Provider', '', '', d.tech.hostingProvider],
    ['CI/CD & Deployment',       '', '', d.tech.cicd],
    [],
    ['  4.2  Existing Integrations'],
    ['Current GDS / Content Provider', '', '', d.integrations.gds],
    ['Existing Booking Engine',        '', '', d.integrations.bookingEngine],
    ['Payment Provider',               '', '', d.integrations.paymentProvider],
    ['Analytics / Attribution',        '', '', d.integrations.analytics],
    ['CRM / Customer Data Platform',   '', '', d.integrations.crm],
    [],
    ['  4.3  White-Label & Branding'],
    ['Brand Colour (Primary)',      '', '', d.branding.primaryColor],
    ['Brand Colour (Secondary)',    '', '', d.branding.secondaryColor],
    ['Primary Font',               '', '', d.branding.primaryFont],
    ['Logo Available?',            '', '', d.branding.logoAvailable],
    ['Custom Domain for Checkout?','', '', d.branding.customDomain],
    ['Language Requirements',      '', '', d.branding.languages],
    ['Currency Display',           '', '', d.branding.currencyDisplay],
    [],
    ['  4.4  Data & Compliance'],
    ['GDPR / Data Residency Requirements', '', '', d.compliance.gdpr],
    ['Data Sharing Restrictions',          '', '', d.compliance.dataSharing],
    ['PII Handling Preferences',           '', '', d.compliance.piiHandling],
    ['Test Environment Available?',        '', '', d.compliance.testEnv],
    ['Sandbox Testing Period Needed',      '', '', d.compliance.sandboxPeriod],
    ['Additional Compliance Notes',        '', '', d.compliance.additionalNotes],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s4), '4 — Technical Details');

  return xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// ─── Build Excel: HolidayHeroes ───────────────────────────────────────────────
function buildHolidayHeroesExcel(d) {
  const wb = xlsx.utils.book_new();

  // Sheet 1: Instructions
  const s0 = [
    ['MOVE — HOLIDAYHEROES PARTNER ONBOARDING QUESTIONNAIRE'],
    ['HolidayHeroes Platform  ·  Powered by Move'],
    [],
    ['Completed via Move Partner Onboarding portal on', new Date().toLocaleDateString('en-GB')],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s0), 'Instructions');

  // Sheet 2: Company & Contacts
  const s1 = [
    ['COMPANY & CONTACT INFORMATION'],
    [],
    ['Company Name',  d.company?.name || ''],
    ['Website',       d.company?.website || ''],
    ['Country / HQ',  d.company?.country || ''],
    ['Company Type',  d.company?.type || ''],
    [],
    ['COMMERCIAL CONTACT'],
    ['Full Name',  d.commercial?.name || ''],
    ['Job Title',  d.commercial?.title || ''],
    ['Email',      d.commercial?.email || ''],
    ['Phone',      d.commercial?.phone || ''],
    [],
    ['TECHNICAL CONTACT'],
    ['Full Name',  d.technical?.name || ''],
    ['Job Title',  d.technical?.title || ''],
    ['Email',      d.technical?.email || ''],
    ['Phone',      d.technical?.phone || ''],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s1), '1 — Company & Contacts');

  // Sheet 3: Market & Routes (Q1-Q7)
  const routeRows = (d.routes || [])
    .filter(r => r.name)
    .map((r, i) => [i + 1, r.name, r.domestic ? 'Yes' : '', r.international ? 'Yes' : '']);

  const s2 = [
    ['MARKET & ROUTES'],
    [],
    ['Q1 — Operating Market', d.market || ''],
    [],
    ['Q2 — Selling Currencies', d.currencies || ''],
    [],
    ['Q3 — Departure Airports', d.departureAirports || ''],
    [],
    ['Q4 — Key Destinations (80% of revenue)'],
    [d.destinations || ''],
    [],
    ['Q5 — Top 20 Routes'],
    ['#', 'Route / Destination', 'Domestic', 'International'],
    ...routeRows,
    [],
    ['Q6 — Total Routes Currently Selling', d.totalRoutes || ''],
    [],
    ['Q7 — % of Routes Making 80% Revenue', d.revenueRoutePct || ''],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s2), '2 — Market & Routes');

  // Sheet 4: Content (Q8-Q10)
  const airRows = (d.airContent || []).filter(a => a.airline)
    .map(a => [a.airline, a.gds, a.tf, a.aggregator, a.direct, a.ndc, a.negotiated]);
  const hotelRows = (d.hotelContent || []).filter(h => h.name)
    .map(h => [h.name, h.gds, h.bedbank, h.directConnect, h.aggregator, h.other]);
  const nonAirRows = (d.nonAirContent || []).filter(n => n.type)
    .map(n => [n.type, n.source, n.negotiated]);

  const s3 = [
    ['CONTENT SOURCES'],
    [],
    ['Q8 — Air Content'],
    ['Airline', 'GDS', 'TF', 'Aggregator', 'Direct', 'NDC', 'Negotiated Pricing'],
    ...airRows,
    [],
    ['Q9 — Hotel Content'],
    ['Name / Group', 'GDS', 'Bedbank(s)', 'Direct Connect', 'Aggregator / Channel Mgr', 'Other'],
    ...hotelRows,
    [],
    ['Q10 — Non-Air Content'],
    ['Type', 'Source / Connection', 'Negotiated Pricing'],
    ...nonAirRows,
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s3), '3 — Content');

  // Sheet 5: Operations & Technology (Q11-Q23)
  const tourOpRows = (d.tourOperators || []).filter(t => t.name)
    .map(t => [t.name, t.connection]);

  const s4 = [
    ['OPERATIONS & TECHNOLOGY'],
    [],
    ['Q11 — Resell Tour Operators?', d.resellTourOps || ''],
    ...(tourOpRows.length ? [['Tour Operator', 'Connection'], ...tourOpRows, []] : [[]]),
    ['Q12 — Payment Gateway', d.paymentGateway || ''],
    [],
    ['Q13 — Paying Party', d.payingParty || ''],
    [],
    ['Q14 — Payment Terms', d.paymentTerms || ''],
    [],
    ['Q15 — Back Office Integrations', d.backOffice || ''],
    [],
    ['Q16 — Current Applications'],
    ['B2B', d.currentApps?.b2b || ''],
    ['B2C', d.currentApps?.b2c || ''],
    ['B2B2C', d.currentApps?.b2b2c || ''],
    [],
    ['Q17 — Future Applications'],
    ['B2B', d.futureApps?.b2b || ''],
    ['B2C', d.futureApps?.b2c || ''],
    ['B2B2C', d.futureApps?.b2b2c || ''],
    [],
    ['Q18 — Replacing Applications'],
    ['B2B', d.replacingApps?.b2b || ''],
    ['B2C', d.replacingApps?.b2c || ''],
    ['B2B2C', d.replacingApps?.b2b2c || ''],
    [],
    ['Q19 — B2B2C Partner Companies', d.b2b2cPartners || ''],
    [],
    ['Q20 — Merchant of Record', d.merchantOfRecord || ''],
    [],
    ['Q21 — B2B Agents'],
    ['Nr of Sub Agents', d.b2bAgents?.count || ''],
    ['% Producing 50+ bookings/yr', d.b2bAgents?.pct || ''],
    [],
    ['Q22 — UI/UX Requirements', d.uiRequirements || ''],
    [],
    ['Q23 — Other Technical Info', d.otherInfo || ''],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s4), '4 — Operations & Tech');

  return xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// ─── POST /submit/ratecore ────────────────────────────────────────────────────
app.post('/submit/ratecore', async (req, res) => {
  try {
    const data = req.body;
    const partnerName = data.company?.name || 'Unknown';
    const buffer = buildRatecoreExcel(data);
    const { filename } = saveSubmission('ratecore', partnerName, data, buffer);
    await sendEmail(partnerName, 'Ratecore', buffer, filename);
    res.json({ ok: true, filename });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ─── POST /submit/holidayheroes ───────────────────────────────────────────────
app.post('/submit/holidayheroes', async (req, res) => {
  try {
    const data = req.body;
    const partnerName = data.company?.name || 'Unknown';
    const buffer = buildHolidayHeroesExcel(data);
    const { filename } = saveSubmission('holidayheroes', partnerName, data, buffer);
    await sendEmail(partnerName, 'HolidayHeroes', buffer, filename);
    res.json({ ok: true, filename });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ─── Backward compat: POST /submit → ratecore ─────────────────────────────────
app.post('/submit', async (req, res) => {
  try {
    const data = req.body;
    const partnerName = data.company?.name || 'Unknown';
    const buffer = buildRatecoreExcel(data);
    const { filename } = saveSubmission('ratecore', partnerName, data, buffer);
    await sendEmail(partnerName, 'Ratecore', buffer, filename);
    res.json({ ok: true, filename });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ─── GET /admin/submissions ───────────────────────────────────────────────────
app.get('/admin/submissions', (req, res) => {
  try {
    const files = fs.readdirSync(SUBMISSIONS_DIR)
      .filter(f => f.endsWith('.json'))
      .sort()
      .reverse();

    const submissions = files.map(f => {
      try {
        return JSON.parse(fs.readFileSync(path.join(SUBMISSIONS_DIR, f), 'utf8'));
      } catch {
        return null;
      }
    }).filter(Boolean);

    res.json(submissions);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Move Onboarding server running on port ${PORT}`));
