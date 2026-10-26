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
        return String(fv).trim();
      }
    }
  }
  return '';
}

const name = pick('full name', 'applicant name', 'name') || 'Guest Tenant';
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

const rentMap = { KE1: 639, KE2: 649, KE3: 619, KE4: 569 };
const depositMap = { KE1: 1278, KE2: 1298, KE3: 1238, KE4: 1138 };
const roomTitles = {
  KE1: 'Master Bedroom (21 m²)',
  KE2: 'Private Balcony Suite (24 m²)',
  KE3: 'Quiet Garden View (12 m²)',
  KE4: 'Cozy Calm Room (11 m²)'
};

const rentAmount = rentMap[room] || 649;
const depositAmount = depositMap[room] || 1298;
const roomTitle = roomTitles[room] || 'Coliving Bedroom';

const emailSubject = `Homelyo Berlin — Booking Request for Room ${room} (${roomTitle.split('(')[0].trim()}) & Document Verification`;
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
• Warm Rent: €${rentAmount} / month (all inclusive: heating, electricity, water, high-speed WiFi)
• Deposit (Kaution): €${depositAmount} (held in escrow in accordance with German tenancy law)
• Move-in Date: ${startDate}
• Initial Duration: ${durationMonths} months

Once we receive and review your documents, we will share the digital tenancy agreement for your review and electronic signature.

If you have any questions about the room, public transit connections, or the move-in process, please feel free to reach out.

Best regards,

The Homelyo Team
Homelyo Coliving Berlin
Email: support@homelyo.eu
Web: https://homelyo.eu`;

const markdownDescription = `### 📋 Homelyo Inbound Booking Application Dossier

* **Applicant Name:** ${name}
* **Email:** ${email}
* **Phone / WhatsApp:** ${phone}
* **Preferred Room:** Room ${room} (${roomTitle})
* **Target Move-In Date:** ${startDate}
* **Stay Duration:** ${durationMonths} months
* **Monthly Warm Rent:** €${rentAmount} / mo (all inclusive)
* **Deposit (Kaution):** €${depositAmount} (held in escrow)
* **Occupation / University:** ${occupation}
* **Current Address:** ${currentAddress}
* **Number of Occupants:** ${occupants}
* **Referral Source:** ${hearAbout}
* **Notes & Inquiries:** ${notes}

---
### ✅ Next Steps & Verification Checklist
- [ ] Identity & Passport Verification
- [ ] German Visa / Residence Permit Check
- [ ] University Enrollment / Employment Contract Check
- [ ] 3 Months Proof of Funds / Stipend Verification
- [ ] Send Prepared Document Request Email Below
- [ ] Prepare Dual-Language Mietvertrag & Escrow Agreement

---
### ✉️ Next-Step Tenant Document Request Email (Draft Prepared)
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

✉️ *A document verification draft email has been prepared in the ClickUp ticket.*`;

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
    markdown_description: markdownDescription,
    chat_alert: chatAlert,
    source: 'Homelyo Booking Engine (Fillout / Website)'
  }
}];
