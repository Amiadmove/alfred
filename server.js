const express = require('express');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

// ─── Config ───────────────────────────────────────────────────────────────────
const NOTIFY_EMAIL   = 'amiad@alfredtravel.io';   // שלח אלי
const SMTP_HOST      = process.env.SMTP_HOST  || 'smtp.gmail.com';
const SMTP_PORT      = process.env.SMTP_PORT  || 587;
const SMTP_USER      = process.env.SMTP_USER;      // set in env
const SMTP_PASS      = process.env.SMTP_PASS;      // set in env
const PORT           = process.env.PORT       || 3000;
const SUBMISSIONS_DIR = path.join(__dirname, 'submissions');

if (!fs.existsSync(SUBMISSIONS_DIR)) fs.mkdirSync(SUBMISSIONS_DIR);

// ─── Build Excel from form data ───────────────────────────────────────────────
function buildExcel(d) {
  const wb = xlsx.utils.book_new();

  // Sheet 1: Instructions
  const s0 = [
    ['ALFRED TRAVEL — ONBOARDING QUESTIONNAIRE'],
    ['Agentic Commerce Integration  ·  Powered by HolidayHeroes x Move'],
    [],
    ['Completed via Alfred Partner Onboarding portal on', new Date().toLocaleDateString('en-GB')],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s0), 'Instructions');

  // Sheet 2: Company & Contacts
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

  // Sheet 3: Volume Forecast
  const monthNames = Array.from({ length: 12 }, (_, i) => `Month ${i + 1}`);
  const totalGBV = d.months.reduce((sum, m) => sum + ((Number(m.bookings) || 0) * (Number(m.avgValue) || 0)), 0);
  const sumField = (f) => d.months.reduce((s, m) => s + (Number(m[f]) || 0), 0);

  const s2 = [
    ['TRAFFIC & BOOKING VOLUME FORECAST — 12 MONTHS POST GO-LIVE'],
    [],
    ['Month', 'Monthly App Sessions', 'Travel Searches', 'Confirmed Bookings', 'Avg. Booking Value (EUR)', 'Expected GBV (EUR)', 'Cancellation Rate (%)', 'Notes'],
    ...d.months.map((m, i) => [
      monthNames[i],
      m.sessions || '',
      m.searches || '',
      m.bookings || '',
      m.avgValue || '',
      (m.bookings && m.avgValue) ? Number(m.bookings) * Number(m.avgValue) : '',
      m.cancellationRate || '',
      m.notes || '',
    ]),
    ['TOTAL (12 mo.)', sumField('sessions'), sumField('searches'), sumField('bookings'), '', totalGBV, '', ''],
    [],
    ['  Additional Context'],
    ['Peak Travel Months',        '', d.forecast.peakMonths],
    ['Seasonal Patterns',         '', d.forecast.seasonalPatterns],
    ['Avg. Lead Time to Departure','', d.forecast.avgLeadTime],
    ['Average Trip Duration',     '', d.forecast.avgTripDuration],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s2), '2 — Volume Forecast');

  // Sheet 4: Routes & Markets
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
    ['Primary User Segments', '', '', d.segments.primarySegments],
    ['Typical Group Size',    '', '', d.segments.groupSize],
    ['Price Sensitivity',     '', '', d.segments.priceSensitivity],
    ['Booking Preference',    '', '', d.segments.bookingPreference],
    ['Avg. Days to Departure','', '', d.segments.avgDaysToDeparture],
    ['Mobile vs Desktop Split','','', d.segments.mobileDesktopSplit],
  ];
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(s3), '3 — Routes & Markets');

  // Sheet 5: Technical Details
  const s4 = [
    ['TECHNICAL INTEGRATION DETAILS'],
    [],
    ['  4.1  Tech Stack'],
    ['Frontend Framework',      '', '', d.tech.frontendFramework],
    ['Backend Language',        '', '', d.tech.backendLanguage],
    ['API Preference',          '', '', d.tech.apiPreference],
    ['Authentication Method',   '', '', d.tech.authMethod],
    ['Hosting / Cloud Provider','', '', d.tech.hostingProvider],
    ['CI/CD & Deployment',      '', '', d.tech.cicd],
    [],
    ['  4.2  Existing Integrations'],
    ['Current GDS / Content Provider', '', '', d.integrations.gds],
    ['Existing Booking Engine',        '', '', d.integrations.bookingEngine],
    ['Payment Provider',               '', '', d.integrations.paymentProvider],
    ['Analytics / Attribution',        '', '', d.integrations.analytics],
    ['CRM / Customer Data Platform',   '', '', d.integrations.crm],
    [],
    ['  4.3  White-Label & Branding'],
    ['Brand Colour (Primary)',   '', '', d.branding.primaryColor],
    ['Brand Colour (Secondary)', '', '', d.branding.secondaryColor],
    ['Primary Font',             '', '', d.branding.primaryFont],
    ['Logo Available?',          '', '', d.branding.logoAvailable],
    ['Custom Domain for Checkout?', '', '', d.branding.customDomain],
    ['Language Requirements',    '', '', d.branding.languages],
    ['Currency Display',         '', '', d.branding.currencyDisplay],
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

// ─── POST /submit ──────────────────────────────────────────────────────────────
app.post('/submit', async (req, res) => {
  try {
    const data = req.body;
    const partnerName = (data.company?.name || 'Unknown').replace(/[^a-z0-9]/gi, '_');
    const timestamp   = new Date().toISOString().slice(0, 10);
    const filename    = `${partnerName}_${timestamp}.xlsx`;
    const filepath    = path.join(SUBMISSIONS_DIR, filename);

    // 1. Generate & save Excel
    const buffer = buildExcel(data);
    fs.writeFileSync(filepath, buffer);
    console.log(`Saved: ${filepath}`);

    // 2. Send email notification (if SMTP configured)
    if (SMTP_USER && SMTP_PASS) {
      const transporter = nodemailer.createTransport({
        host: SMTP_HOST,
        port: SMTP_PORT,
        secure: false,
        auth: { user: SMTP_USER, pass: SMTP_PASS },
      });

      await transporter.sendMail({
        from: `"Alfred Onboarding" <${SMTP_USER}>`,
        to: NOTIFY_EMAIL,
        subject: `New Onboarding Submission — ${data.company?.name || 'Unknown Partner'}`,
        text: `A new partner onboarding form has been submitted.\n\nPartner: ${data.company?.name}\nContact: ${data.commercial?.name} (${data.commercial?.email})\n\nThe completed Excel file is attached.`,
        attachments: [{ filename, content: buffer }],
      });
      console.log(`Email sent to ${NOTIFY_EMAIL}`);
    }

    res.json({ ok: true, filename });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.listen(PORT, () => console.log(`Alfred Onboarding server running on port ${PORT}`));
