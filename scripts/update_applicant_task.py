import urllib.request
import json

TOKEN = 'pk_105682594_CCTOZ8K23KHPSAY3JQTVV1G23QTN63IG'
HEADERS = {'Authorization': TOKEN, 'Content-Type': 'application/json'}

desc = """### 📋 Homelyo Inbound Booking Request & Tenant Onboarding Dossier

* **Applicant Name:** Ajithkumar Velasseril Sreenivasan
* **Email:** ajithkmvs123@gmail.com
* **Phone / WhatsApp:** +918606817191
* **Selected Room:** Room KE2 (Private Balcony Suite — 24 m²)
* **Target Move-In Date:** 2026-11-02
* **Stay Duration:** 6 months (2026-11-02 → 2027-04-30)
* **Monthly Warm Rent:** €649.00 / month (all inclusive)
* **Deposit (Kaution):** €1,298.00 (held in escrow)
* **Occupation / University:** Postdoctoral researcher at Freie Universität Berlin (FU Berlin)
* **Notes & Inquiries:** None

---
### 📁 Generated Legal Documents & Assets (Google Drive)
* 📂 [Open Tenant Google Drive Folder](https://drive.google.com/drive/u/0/folders/1WMUDX6oqFx_7HbJSJGUes4m6Rwb_C86z)
* 📝 [Open Dual-Language Rental Contract (Google Doc)](https://docs.google.com/document/d/1mTkoik1ok_RhR9FnBe4gWjrn3KEntrPAFT5VHPBOAH8/edit)
* 📑 [Download PDF Rental Contract](https://drive.google.com/file/d/1Y7F8_-VU6jsP1bZADK65bqjGxQKIlfXx/view)
* 📊 Google Sheet: `tenant_details` & `rent_log` synchronized successfully
* 🏠 Related Active Lease Task: [Task #1amhxucnb5c](https://app.clickup.com/t/1amhxucnb5c)

---
### ✉️ Next-Step Tenant Document Request Email (Draft Prepared)
**To:** ajithkmvs123@gmail.com  
**From:** Homelyo Team <support@homelyo.eu>  
**Subject:** Homelyo Berlin — Booking Request for Room KE2 (Balcony Suite) & Document Verification  

```text
Dear Ajithkumar,

Thank you for your interest in Homelyo Berlin! We have received your booking application for Room KE2 (Balcony Suite) with your requested move-in date on November 2, 2026.

Congratulations on your postdoctoral appointment at Freie Universität Berlin (FU Berlin). We would be delighted to welcome you to our coliving community.

To proceed with verifying your booking and preparing your digital tenancy agreement (Mietvertrag & Wohnungsgeberbestätigung for your Bürgeramt city registration), please reply to this email with copies of the following documents:

1. Identification & Legal Stay:
   • Clear scan or photo of your valid Passport (photo & biographical page).
   • Valid German Visa / Entry Visa (Type D) or German Residence Permit (Aufenthaltstitel).

2. University & Academic Appointment:
   • Official FU Berlin Postdoctoral Appointment / Fellowship Award Letter or Academic Employment Contract.

3. Proof of Financial Capability:
   • Bank statements from the last 3 months, or official stipend / salary disbursement documentation confirming sufficient funds.

4. Emergency Contact Information:
   • Full name, relationship, contact phone number, and email address of your designated emergency contact.

Room Summary & Financial Overview:
• Room: KE2 (Balcony Suite)
• Warm Rent: €649 / month (all inclusive: heating, electricity, water, high-speed WiFi)
• Deposit (Kaution): €1,298 (held in escrow in accordance with German tenancy law)
• Move-in Date: November 2, 2026
• Initial Duration: 6 months

Once we receive and review your documents, we will share the digital tenancy agreement for your review and electronic signature.

Best regards,

The Homelyo Team
Homelyo Coliving Berlin
Email: support@homelyo.eu
Web: https://homelyo.eu
```
"""

payload = {
    'markdown_description': desc
}

req = urllib.request.Request(
    'https://api.clickup.com/api/v2/task/1ampw4tj1tz',
    data=json.dumps(payload).encode('utf-8'),
    headers=HEADERS,
    method='PUT'
)
with urllib.request.urlopen(req) as resp:
    print('Updated task 1ampw4tj1tz status:', resp.status)

comment_payload = {
    'comment_text': '✉️ Prepared Email Draft for Document Vetting added to task description. Google Drive folder, dual-language contract (Google Doc & PDF), and Google Sheet have been generated.',
    'notify_all': False
}
req2 = urllib.request.Request(
    'https://api.clickup.com/api/v2/task/1ampw4tj1tz/comment',
    data=json.dumps(comment_payload).encode('utf-8'),
    headers=HEADERS,
    method='POST'
)
with urllib.request.urlopen(req2) as resp2:
    print('Comment posted status:', resp2.status)
