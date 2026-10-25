import urllib.request
import json

TOKEN = "pk_308598557_9CP65EANIWJ0XHF9NUOFGMFUDCAMPPM5"
LIST_ID = "1500350000002188" # Dashboards, Looker & Grafana Hub
HEADERS = {"Authorization": TOKEN, "Content-Type": "application/json"}

content = """# 🏛️ Homelyo Unified System Architecture & Whiteboard Blueprint
**Property:** Lichtenrader Damm 80, 12305 Berlin (Co-Living 4-Room Property)
**Owner / Management Hub:** ClickUp Space 'Homelyo' (ID: 1500350000000761)

---

## 🎨 Whiteboard Visual Architecture Map

```
+---------------------------------------------------------------------------------------------------+
|                                  HOMELYO PUBLIC CLIENT LAYER                                      |
|                                     https://www.homelyo.eu/                                       |
|  +--------------------+  +--------------------+  +--------------------+  +---------------------+  |
|  |   Main Portal      |  |   Room Showcases   |  |   Amenities Tour   |  |   Move-in Package   |  |
|  |   index.html       |  |   ke1 to ke4.html  |  |   amenities.html   |  | move-in-package.html|  |
|  +--------------------+  +--------------------+  +--------------------+  +---------------------+  |
|          |                         |                        |                       |             |
|          v                         v                        v                       v             |
|  +--------------------------------------------+  +---------------------------------------------+  |
|  |           Sara AI Concierge Widget         |  |         Lead & Contact Capture Forms        |  |
|  |  assistant/chat-widget.js & index.html     |  |       book.html (Fillout/ClickUp Form)      |  |
|  |     (Avatar, speech warmup, quick FAQ)     |  |       contact.html (Interactive Form)       |  |
|  +--------------------------------------------+  +---------------------------------------------+  |
+---------------------------------------------------------------------------------------------------+
                                   |                                              |
                          POST /message (JSON)                         POST /task (Lead Payload)
                                   v                                              |
+-------------------------------------------------------------+                   |
|                   SERVERLESS PROXY LAYER                    |                   |
|       Cloudflare Worker: assistant-proxy.kylashpriya        |                   |
|      - CORS Policy & Origin Filtering                       |                   |
|      - Session Management & Intent Parsing                  |                   |
|      - LLM Inference & Action Dispatcher                    |                   |
|      - High-Intent Lead Webhook Trigger                     |                   |
+-------------------------------------------------------------+                   |
                                   |                                              |
                                   +----------------------+-----------------------+
                                                          |
                                           Direct HTTPS REST API v2
                                           (Authorization: Bearer Token)
                                                          v
+---------------------------------------------------------------------------------------------------+
|                                CLICKUP COMMAND CENTER (HOMELYO SPACE)                             |
|                                       Space ID: 1500350000000761                                  |
|                                                                                                   |
|  [Folder 1: Property & Tenant Management]                                                         |
|  • List: Tenants & Room Occupancy (KE1-KE4, rent, deposits, leases, Anmeldung, protocols)        |
|  • List: Booking Requests & Tenant Inquiries (Kanban pipeline, vetting, move-in calendar)        |
|                                                                                                   |
|  [Folder 2: Financials & Operating Expenses]                                                      |
|  • List: Operating Expenses & Utility Bills (Hausgeld, Vattenfall, Vodafone, GEZ, Grundsteuer)   |
|  • List: Rent & Deposit Ledger (Monthly collections, rent roll, escrow deposit security)          |
|                                                                                                   |
|  [Folder 3: Amenities & Asset Inventory]                                                          |
|  • List: Apartment Inventory & Assets (Samsung fridge, 4-mode shower, induction, furniture)      |
|  • List: Maintenance & Supplies Restock (Work orders, contractor tickets, routine upkeep)        |
|                                                                                                   |
|  [Folder 4: Sara AI Assistant & Enquiries]                                                        |
|  • List: Sara Chat Conversations & Leads (Transcripts, sentiment, room interest, follow-ups)     |
|  • List: Website Contact & Support Inquiries (Inbound web/email inquiries, SLA tracking)         |
|                                                                                                   |
|  [Folder 5: Dashboards, Analytics & Architecture Hub]                                             |
|  • List: Dashboards, Looker & Grafana Hub (Consolidated master index of all external systems)    |
|                                                                                                   |
|  Space Views: Executive Dashboard | Whiteboard | Live Grafana Embed | 3x Looker Embeds | GDrive  |
+---------------------------------------------------------------------------------------------------+
                                   |                                              |
                     Bi-Directional Data Sync                       Metrics & Storage References
                                   v                                              v
+-------------------------------------------------------------+  +----------------------------------+
|                  BUSINESS INTELLIGENCE LAYER                |  |       CLOUD STORAGE VAULT        |
|                                                             |  |                                  |
|  • Looker Studio 1 (Financial Overview & Accounting):       |  |  • Google Drive Central Folder:  |
|    https://datastudio.google.com/u/0/reporting/49f8c6ad...  |  |    https://drive.google.com/...  |
|  • Looker Studio 2 (Revenue Streams & Billings):            |  |    - Deeds & Grundbuchauszug     |
|    https://datastudio.google.com/u/0/reporting/469f33c1...  |  |    - Signed Leases (Mietverträge)|
|  • Looker Studio 3 (Occupancy & Operational KPIs):          |  |    - Handover Protocols          |
|    https://datastudio.google.com/u/0/reporting/012f2dfd...  |  |    - Hausgeld Abrechnungen       |
|  • Grafana Cloud (Live Telemetry & Infra Monitoring):       |  |    - Renovation & Tax Invoices   |
|    https://kylash.grafana.net/dashboards/f/homelyo-...      |  +----------------------------------+
+-------------------------------------------------------------+
```

---

## 🔄 Automated Data Flow & Event Cycles

1. **Lead Generation & Booking Request**:
   - Prospective tenant submits request via book.html, contact.html, or Sara Chat.
   - Task created in 'Booking Requests' or 'Sara Chat Conversations'.
   - Notification sent to owner; status set to 'Document Vetting'.

2. **Lease Execution & Onboarding**:
   - Contract signed (English/German); security deposit escrow confirmed.
   - Lead moved to 'Tenants & Room Occupancy'.
   - Move-in package dispatched; Anmeldung confirmation delivered.

3. **Financial Reconciliation**:
   - 1st of every month: Rent collection checked across KE1-KE4 in 'Rent Ledger'.
   - Recurring Hausgeld, electricity, internet auto-logged in 'Operating Expenses'.
   - Data feeds into Looker Studio Report 1 & 2 for net yield calculation.

4. **Maintenance & Turnaround**:
   - Tenant flags an issue via Maintenance Form.
   - Task triaged by urgency (Emergency, High, Normal, Scheduled).
   - Handyman assigned; cost recorded; resolved with sign-off.
"""

payload = {
    "name": "🏛️ Technical Architecture & Whiteboard Blueprint (Complete)",
    "description": content,
    "priority": 1
}

url = f"https://api.clickup.com/api/v2/list/{LIST_ID}/task"
req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=HEADERS)
with urllib.request.urlopen(req) as resp:
    data = json.loads(resp.read().decode())
    print("Architecture Blueprint Task Created:", data.get("id"), data.get("name"))
