const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const { put } = require('@vercel/blob');

const NOTIFY_EMAIL = process.env.NOTIFY_EMAIL || 'amiad@alfredtravel.io';

function buildExcel(d) {
  const wb = xlsx.utils.book_new();

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['MOVE — RATECORE PARTNER ONBOARDING QUESTIONNAIRE'],
    ['Ratecore Integration  ·  Powered by Move'],
    [],
    ['Completed on', new Date().toLocaleDateString('en-GB')],
  ]), 'Instructions');

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['COMPANY & CONTACT INFORMATION'],
    [],
    ['Company Name',               d.company?.name           || ''],
    ['Website',                    d.company?.website        || ''],
    ['Registered Address',         d.company?.address        || ''],
    ['Company Reg. Number',        d.company?.regNumber      || ''],
    ['VAT Number',                 d.company?.vatNumber      || ''],
    ['Primary Market(s)',          d.company?.primaryMarkets || ''],
    ['App Platform(s)',            d.company?.appPlatforms   || ''],
    [],
    ['COMMERCIAL CONTACT'],
    ['Full Name',  d.commercial?.name  || ''],
    ['Job Title',  d.commercial?.title || ''],
    ['Email',      d.commercial?.email || ''],
    ['Phone',      d.commercial?.phone || ''],
    [],
    ['TECHNICAL CONTACT'],
    ['Full Name',              d.technical?.name   || ''],
    ['Job Title',              d.technical?.title  || ''],
    ['Email',                  d.technical?.email  || ''],
    ['Phone',                  d.technical?.phone  || ''],
    ['GitHub / GitLab handle', d.technical?.github || ''],
    [],
    ['FINANCE / BILLING CONTACT'],
    ['Full Name',                  d.finance?.name       || ''],
    ['Email',                      d.finance?.email      || ''],
    ['Preferred Invoice Currency', d.finance?.currency   || ''],
    ['PO Number Required?',        d.finance?.poRequired || ''],
  ]), '1 — Company & Contacts');

  const monthNames = Array.from({length:12}, (_,i) => `Month ${i+1}`);
  const totalGBV = (d.months||[]).reduce((s,m) => s+(Number(m.bookings)||0)*(Number(m.avgValue)||0), 0);
  const sumF = f => (d.months||[]).reduce((s,m) => s+(Number(m[f])||0), 0);

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['TRAFFIC & BOOKING VOLUME FORECAST — 12 MONTHS'],
    [],
    ['Month','App Sessions','Travel Searches','Confirmed Bookings','Avg. Booking Value (EUR)','Expected GBV (EUR)','Cancel Rate (%)','Notes'],
    ...(d.months||[]).map((m,i) => [monthNames[i], m.sessions||'', m.searches||'', m.bookings||'', m.avgValue||'',
      (m.bookings&&m.avgValue)?Number(m.bookings)*Number(m.avgValue):'', m.cancellationRate||'', m.notes||'']),
    ['TOTAL', sumF('sessions'), sumF('searches'), sumF('bookings'), '', totalGBV, '', ''],
    [],
    ['Peak Travel Months',         d.forecast?.peakMonths       || ''],
    ['Seasonal Patterns',          d.forecast?.seasonalPatterns || ''],
    ['Avg. Lead Time to Departure',d.forecast?.avgLeadTime      || ''],
    ['Average Trip Duration',      d.forecast?.avgTripDuration  || ''],
  ]), '2 — Volume Forecast');

  const airportTotal = (d.airports||[]).reduce((s,a) => s+(Number(a.traffic)||0), 0);
  const destTotal    = (d.destinations||[]).reduce((s,dest) => s+(Number(dest.traffic)||0), 0);

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['ROUTES, DESTINATIONS & USER SEGMENTS'],
    [],
    ['Top Departure Airports'],
    ['#','Airport Name','IATA Code','City / Country','Est. % Traffic','Notes'],
    ...(d.airports||[]).map((a,i) => [i+1, a.name, a.iata, a.cityCountry, a.traffic||'', a.notes]),
    ['Total','','','',airportTotal+'%',''],
    [],
    ['Top Destinations'],
    ['#','City / Region','Country','Product Type','Est. % Traffic','Notes'],
    ...(d.destinations||[]).map((dest,i) => [i+1, dest.cityRegion, dest.country, dest.productType, dest.traffic||'', dest.notes]),
    ['Total','','','',destTotal+'%',''],
    [],
    ['User Segments & Travel Profile'],
    ['Primary User Segments',   d.segments?.primarySegments  || ''],
    ['Typical Group Size',      d.segments?.groupSize        || ''],
    ['Price Sensitivity',       d.segments?.priceSensitivity || ''],
    ['Booking Preference',      d.segments?.bookingPreference|| ''],
    ['Avg. Days to Departure',  d.segments?.avgDaysToDeparture||''],
    ['Mobile vs Desktop Split', d.segments?.mobileDesktopSplit||''],
  ]), '3 — Routes & Markets');

  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet([
    ['TECHNICAL INTEGRATION DETAILS'],
    [],
    ['Tech Stack'],
    ['Frontend Framework',       d.tech?.frontendFramework || ''],
    ['Backend Language',         d.tech?.backendLanguage   || ''],
    ['API Preference',           d.tech?.apiPreference     || ''],
    ['Authentication Method',    d.tech?.authMethod        || ''],
    ['Hosting / Cloud Provider', d.tech?.hostingProvider   || ''],
    ['CI/CD & Deployment',       d.tech?.cicd              || ''],
    [],
    ['Existing Integrations'],
    ['Current GDS / Content Provider', d.integrations?.gds            || ''],
    ['Existing Booking Engine',        d.integrations?.bookingEngine  || ''],
    ['Payment Provider',               d.integrations?.paymentProvider|| ''],
    ['Analytics / Attribution',        d.integrations?.analytics      || ''],
    ['CRM / Customer Data Platform',   d.integrations?.crm            || ''],
    [],
    ['White-Label & Branding'],
    ['Brand Colour (Primary)',   d.branding?.primaryColor   || ''],
    ['Brand Colour (Secondary)', d.branding?.secondaryColor || ''],
    ['Primary Font',             d.branding?.primaryFont    || ''],
    ['Logo Available?',          d.branding?.logoAvailable  || ''],
    ['Custom Domain for Checkout?',d.branding?.customDomain || ''],
    ['Language Requirements',    d.branding?.languages      || ''],
    ['Currency Display',         d.branding?.currencyDisplay|| ''],
    [],
    ['Data & Compliance'],
    ['GDPR / Data Residency',    d.compliance?.gdpr             || ''],
    ['Data Sharing Restrictions',d.compliance?.dataSharing      || ''],
    ['PII Handling',             d.compliance?.piiHandling      || ''],
    ['Test Environment?',        d.compliance?.testEnv          || ''],
    ['Sandbox Period',           d.compliance?.sandboxPeriod    || ''],
    ['Additional Notes',         d.compliance?.additionalNotes  || ''],
  ]), '4 — Technical Details');

  return xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  try {
    const d = req.body;
    const partnerName = d.company?.name || 'Unknown';
    const id = Date.now().toString();
    const buffer = buildExcel(d);
    const filename = `RC_${partnerName.replace(/[^a-z0-9]/gi, '_')}_${new Date().toISOString().slice(0,10)}.xlsx`;

    const submissionData = {
      id,
      platform: 'ratecore',
      partnerName,
      submittedAt: new Date().toISOString(),
      data: d,
    };

    // Save to Vercel Blob (optional — requires BLOB_READ_WRITE_TOKEN)
    if (process.env.BLOB_READ_WRITE_TOKEN) {
      const xlsxBlob = await put(`submissions/${id}.xlsx`, Buffer.from(buffer), {
        access: 'private',
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      submissionData.excelUrl = xlsxBlob.url;
      await put(`submissions/${id}.json`, JSON.stringify(submissionData), {
        access: 'private',
        contentType: 'application/json',
      });
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
        subject: `New Ratecore Onboarding — ${partnerName}`,
        text: `New Ratecore partner onboarding submitted.\n\nPartner: ${partnerName}\nContact: ${d.commercial?.name||'—'} (${d.commercial?.email||'—'})\n\nExcel file attached.`,
        attachments: [{ filename, content: buffer }],
      });
    }

    res.status(200).json({ ok: true, filename });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
};
