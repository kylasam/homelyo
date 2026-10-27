const root = $input.first().json;
const body = root.body || root;

function collectNamedArrays(value, found = []) {
  if (!value || typeof value !== 'object') return found;
  if (Array.isArray(value.questions)) found.push(value.questions);
  if (Array.isArray(value.urlParameters)) found.push(value.urlParameters);
  if (Array.isArray(value.calculations)) found.push(value.calculations);
  if (Array.isArray(value)) {
    for (const item of value) collectNamedArrays(item, found);
  } else {
    for (const item of Object.values(value)) collectNamedArrays(item, found);
  }
  return found;
}

const fields = {};
for (const list of collectNamedArrays(body)) {
  for (const item of list) {
    if (!item || !item.name) continue;
    fields[String(item.name).trim().toLowerCase()] = item.value ?? '';
  }
}
if (Object.keys(fields).length === 0 && typeof body === 'object') {
  for (const [k, v] of Object.entries(body)) {
    fields[String(k).trim().toLowerCase()] = v;
  }
}

function pick(...keys) {
  for (const key of keys) {
    const kLow = key.toLowerCase();
    for (const [fk, fv] of Object.entries(fields)) {
      if (fk.includes(kLow) && fv !== undefined && fv !== null && String(fv).trim() !== '') {
        if (typeof fv === 'object') {
          try {
            const parts = [fv.address, fv.address2, fv.city, fv.state, fv.postalCode, fv.country].filter(Boolean);
            if (parts.length > 0) return parts.join(', ');
            return Object.values(fv).filter(v => typeof v === 'string').join(', ');
          } catch (e) {
            return JSON.stringify(fv);
          }
        }
        return String(fv).trim();
      }
    }
  }
  return '';
}

const name = pick('full name', 'applicant name', 'name') || 'Homelyo Applicant';
const email = pick('email address', 'email') || 'support@homelyo.eu';
const phone = pick('phone number', 'phone', 'whatsapp') || 'Not specified';
const rawRoom = pick('preferred room', 'selected room', 'room preference', 'room') || 'KE2';
const roomMatch = rawRoom.match(/KE[1-4]/i);
const room = roomMatch ? roomMatch[0].toUpperCase() : 'KE2';

const startDate = pick('desired move-in date', 'move-in date (preferred)', 'move-in date', 'target move-in', 'start date', 'start') || '2026-11-01';
const rawDuration = pick('stay duration (months)', 'stay duration', 'length of stay', 'duration', 'months') || '6';
let durationMonths = parseInt(String(rawDuration).replace(/[^\d]/g, '') || '6', 10);
if (isNaN(durationMonths) || durationMonths < 1) durationMonths = 6;

const occupation = pick('current city & occupation / university', 'occupation / university', 'occupation', 'university', 'profession') || 'Not specified';
const currentAddress = pick('current address', 'address', 'city') || 'Not specified';
const occupants = pick('number of occupants', 'occupants') || '1';
const hearAbout = pick('how did you hear about us', 'source') || 'Direct Website';
const notes = pick('notes or questions for owner', 'additional requests or question', 'notes', 'inquiries') || 'None';

const rentMap = { KE1: 640, KE2: 649, KE3: 615, KE4: 569 };
const depositAmount = 1000;
const roomTitles = {
  KE1: 'Master Bedroom',
  KE2: 'Private Balcony Suite',
  KE3: 'Quiet Garden View',
  KE4: 'Cozy Studio Room'
};

const rentAmount = rentMap[room] || 649;
const roomTitle = roomTitles[room] || 'Coliving Suite';

const emailSubject = `Homelyo Berlin — Booking Request for Room ${room} (${roomTitle}) & Document Verification`;

const emailBody = `Dear ${name},

Thank you for your interest in Homelyo Berlin! We have received your booking application for Room ${room} (${roomTitle}) with your requested move-in date on ${startDate}.

We would be delighted to welcome you to our coliving community in Berlin Lichtenrade.

To proceed with verifying your booking and preparing your digital tenancy agreement (Mietvertrag & Wohnungsgeberbestätigung for your Bürgeramt city registration), please reply to this email with copies of the following documents:

1. Identification & Legal Stay:
   • Clear scan or photo of your valid Passport (photo & biographical page).
   • Valid German Visa / Entry Visa (Type D) or German Residence Permit (Aufenthaltstitel).

2. University & Employment Verification:
   • Official University Admission / Academic Appointment Letter or Employment Contract.

3. Proof of Financial Capability:
   • Bank statements from the last 3 months, or official stipend / salary disbursement documentation confirming sufficient funds.

4. Emergency Contact Information:
   • Full name, relationship, contact phone number, and email address of your designated emergency contact.

Room Summary & Financial Overview:
• Room: ${room} (${roomTitle})
• Warm Rent: €${rentAmount} / month (all-inclusive: heating, electricity, water, high-speed Wi-Fi)
• Deposit (Kaution): 1000€ (held in escrow in accordance with German tenancy law)
• Move-in Date: ${startDate}
• Initial Duration: ${durationMonths} months

Once we receive and review your documents, we will share the digital tenancy agreement for your review and electronic signature.

If you have any questions about the room, public transit connections, or the move-in process, please feel free to reach out.

Best regards,

The Homelyo Team
Homelyo Coliving Berlin
Lichtenrader Damm 80, 12305 Berlin
Email: support@homelyo.eu
Web: https://www.homelyo.eu`;

const emailHtml = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background-color: #f8fafc; color: #1e293b; margin: 0; padding: 24px; line-height: 1.6; }
.card { max-width: 620px; margin: 0 auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
.header { background: #0f172a; padding: 24px 32px; border-bottom: 3px solid #00D084; text-align: left; }
.header h1 { margin: 0; color: #ffffff; font-size: 22px; font-weight: 700; }
.header p { margin: 4px 0 0; color: #94a3b8; font-size: 13px; }
.content { padding: 32px; }
h2 { font-size: 16px; color: #0f172a; margin-top: 24px; margin-bottom: 12px; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px; }
ol { padding-left: 20px; margin: 12px 0; }
li { margin-bottom: 12px; }
ul { margin: 6px 0 0; padding-left: 18px; color: #475569; }
.summary-box { background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 18px 22px; margin: 24px 0; }
.summary-box ul { list-style: none; padding-left: 0; margin: 0; }
.summary-box li { margin-bottom: 6px; color: #166534; font-size: 14px; }
.footer { background: #f8fafc; padding: 20px 32px; border-top: 1px solid #e2e8f0; font-size: 13px; color: #64748b; }
a { color: #00D084; text-decoration: none; font-weight: 600; }
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <h1>🏠 Homelyo Coliving Berlin</h1>
    <p>Lichtenrader Damm 80, 12305 Berlin</p>
  </div>
  <div class="content">
    <p>Dear <strong>${name}</strong>,</p>
    <p>Thank you for your interest in Homelyo Berlin! We have received your booking application for <strong>Room ${room} (${roomTitle})</strong> with your requested move-in date on <strong>${startDate}</strong>.</p>
    <p>We would be delighted to welcome you to our coliving community in Berlin Lichtenrade.</p>
    
    <h2>📄 Required Verification Documents</h2>
    <p>To proceed with verifying your booking and preparing your digital tenancy agreement (<em>Mietvertrag &amp; Wohnungsgeberbestätigung</em> for your Bürgeramt city registration), please reply to this email with copies of the following documents:</p>
    
    <ol>
      <li><strong>Identification &amp; Legal Stay:</strong>
        <ul>
          <li>Clear scan or photo of your valid Passport (photo &amp; biographical page).</li>
          <li>Valid German Visa / Entry Visa (Type D) or German Residence Permit (<em>Aufenthaltstitel</em>).</li>
        </ul>
      </li>
      <li><strong>University &amp; Employment Verification:</strong>
        <ul>
          <li>Official University Admission / Academic Appointment Letter or Employment Contract.</li>
        </ul>
      </li>
      <li><strong>Proof of Financial Capability:</strong>
        <ul>
          <li>Bank statements from the last 3 months, or official stipend / salary disbursement documentation confirming sufficient funds.</li>
        </ul>
      </li>
      <li><strong>Emergency Contact Information:</strong>
        <ul>
          <li>Full name, relationship, contact phone number, and email address of your designated emergency contact.</li>
        </ul>
      </li>
    </ol>
    
    <div class="summary-box">
      <strong>Room Summary &amp; Financial Overview:</strong>
      <ul>
        <li>• <strong>Room:</strong> Room ${room} (${roomTitle})</li>
        <li>• <strong>Warm Rent:</strong> €${rentAmount} / month (all-inclusive: heating, electricity, water, high-speed Wi-Fi)</li>
        <li>• <strong>Deposit (Kaution):</strong> 1000€ (held in escrow in accordance with German tenancy law)</li>
        <li>• <strong>Move-in Date:</strong> ${startDate}</li>
        <li>• <strong>Initial Duration:</strong> ${durationMonths} months</li>
      </ul>
    </div>
    
    <p>Once we receive and review your documents, we will share the digital tenancy agreement for your review and electronic signature.</p>
    <p>If you have any questions about the room, public transit connections, or the move-in process, please feel free to reach out.</p>
    
    <p style="margin-top: 24px;">Best regards,<br>
    <strong>The Homelyo Team</strong><br>
    Homelyo Coliving Berlin</p>
  </div>
  <div class="footer">
    📍 Lichtenrader Damm 80, 12305 Berlin<br>
    📧 Email: <a href="mailto:support@homelyo.eu">support@homelyo.eu</a> &nbsp;|&nbsp; 🌐 Web: <a href="https://www.homelyo.eu" target="_blank">https://www.homelyo.eu</a>
  </div>
</div>
</body>
</html>`;

const adminEmailSubject = `📥 [NEW INQUIRY] ${name} — Room ${room} (${startDate})`;

const adminEmailText = `New Inbound Booking Application Received!

Applicant Details:
• Name: ${name}
• Email: ${email}
• Phone: ${phone}
• Room: Room ${room} (${roomTitle})
• Move-In: ${startDate}
• Duration: ${durationMonths} months
• Warm Rent: €${rentAmount} / mo
• Deposit: €${depositAmount}
• Occupation: ${occupation}
• Address: ${currentAddress}
• Occupants: ${occupants}
• Referral Source: ${hearAbout}
• Notes / Questions: ${notes}

A document verification request has been automatically dispatched to ${email}.`;

const adminEmailHtml = `<!DOCTYPE html>
<html>
<body style="font-family: sans-serif; color: #1e293b; padding: 20px;">
  <h2>📥 New Inbound Booking Application Received</h2>
  <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Applicant:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">${name}</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Email:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;"><a href="mailto:${email}">${email}</a></td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Phone:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">${phone}</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Room Preference:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">Room ${room} (${roomTitle})</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Move-In Date:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">${startDate}</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Duration:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">${durationMonths} months</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Warm Rent:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">€${rentAmount} / month</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Deposit:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">1000€</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Occupation / Univ:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">${occupation}</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Current Address:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">${currentAddress}</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Occupants:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">${occupants}</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Referral Source:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">${hearAbout}</td></tr>
    <tr><td style="padding: 8px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Notes / Inquiries:</td><td style="padding: 8px; border-bottom: 1px solid #e2e8f0;">${notes}</td></tr>
  </table>
  <p style="margin-top: 20px; font-size: 13px; color: #64748b;">Automated document verification request has been dispatched to ${email}.</p>
</body>
</html>`;

const markdownDescription = `### 📋 Homelyo Inbound Booking Application Dossier

* **Applicant Name:** ${name}
* **Email:** ${email}
* **Phone / WhatsApp:** ${phone}
* **Preferred Room:** Room ${room} (${roomTitle})
* **Target Move-In Date:** ${startDate}
* **Stay Duration:** ${durationMonths} months
* **Monthly Warm Rent:** €${rentAmount} / mo (all inclusive)
* **Deposit (Kaution):** €1000 (held in escrow)
* **Occupation / University:** ${occupation}
* **Current Address:** ${currentAddress}
* **Number of Occupants:** ${occupants}
* **Referral Source:** ${hearAbout}
* **Notes & Inquiries:** ${notes}

---
### ✅ Next Steps & Verification Checklist
- [ ] Identity & Passport Verification
- [ ] German Visa / Residence Permit Check
- [ ] 3 Months Bank Statements / Proof of Funds
- [x] Auto-Sent Document Request Email to Tenant (${email})
- [ ] Prepare Dual-Language Mietvertrag & Escrow Agreement

---
### ✉️ Document Request Email Sent
**To:** ${email}  
**From:** Homelyo Team <support@homelyo.eu>  
**Subject:** ${emailSubject}  

` + '```text\n' + emailBody + '\n```';

const chatAlert = `📋 **[NEW HOMELYO BOOKING REQUEST]**
• **Applicant:** ${name}
• **Email:** ${email} | **Phone:** ${phone}
• **Room:** ${room} (${roomTitle}) — €${rentAmount}/mo
• **Move-In:** ${startDate} (${durationMonths} months)
• **Occupation:** ${occupation}
• **Notes:** ${notes}

✉️ *Document verification request auto-dispatched to applicant (${email}) and notification to support@homelyo.eu.*`;

return [{
  json: {
    name,
    email,
    phone,
    room,
    room_title: roomTitle,
    start_date: startDate,
    duration_months: durationMonths,
    rent_amount: rentAmount,
    deposit_amount: depositAmount,
    occupation,
    current_address: currentAddress,
    occupants,
    hear_about: hearAbout,
    notes,
    email_subject: emailSubject,
    email_body: emailBody,
    email_html: emailHtml,
    admin_email_subject: adminEmailSubject,
    admin_email_text: adminEmailText,
    admin_email_html: adminEmailHtml,
    markdown_description: markdownDescription,
    chat_alert: chatAlert,
    source: 'Homelyo Booking Engine (Fillout / Website)'
  }
}];
