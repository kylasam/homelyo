import urllib.request
import json

TOKEN = "pk_308598557_9CP65EANIWJ0XHF9NUOFGMFUDCAMPPM5"
HEADERS = {"Authorization": TOKEN, "Content-Type": "application/json"}

cockpit_md = """# 🚀 Homelyo Executive Cockpit — Quick Glance

### 📍 Property Overview
* **Address:** Lichtenrader Damm 80, 12305 Berlin-Lichtenrade
* **Total Property Size:** 96 m² | 4 Private Bedrooms + Shared Kitchen + Renovated Luxury Shower + Dedicated Common Luggage Storage Room
* **Target Tenancy:** Professionals, Master Students, and Expats in Berlin

---

### 🛏️ Live Room Occupancy & Free Dates (Accurate as of Oct 2026)
| Room | Feature | Size | Warm Rent (All-inclusive) | Current Status | Available From |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **KE1** | Master Bedroom | 21 m² | **€629 / mo** | 🟢 **FREE / VACANT** | **Nov 1, 2026** |
| **KE2** | Private Balcony Suite | ~18 m² | **€649 / mo** | 🟢 **FREE / VACANT** | **Nov 1, 2026** |
| **KE3** | Quiet Garden View | ~16 m² | **€619 / mo** | 🟡 Occupied (Lease ends Mar 31) | **Apr 1, 2027** |
| **KE4** | Cozy Calm (Thick Mattress) | ~14 m² | **€569 / mo** | 🟡 Occupied (Lease ends Jul 31) | **Aug 1, 2027** |

---

### 📅 Live Occupancy Timeline & Calendars
* 👉 [Open Timeline (Gantt) View](https://app.clickup.com/90122087374/v/tl/2kxv3uye-2212)
* 👉 [Open Lease Calendar View](https://app.clickup.com/90122087374/v/l/1500350000002180?t=2kxv3uye-1632)
* 👉 [Open Tenants Table View](https://app.clickup.com/90122087374/v/li/1500350000002180?t=2kxv3uye-1572)

---

### 💰 Financial Rollup, Escrow & Tenant Deposit Settlements
* **Gross Potential Monthly Revenue (100% Occupancy):** €2,466.00 / month (€29,592.00 / year)
* **Total Operating Expenses (All Utilities + Hausgeld):** €617.85 / month
  * Hausgeld (HOA, Reserve, Maintenance): €385.00 / mo
  * Electricity (Vattenfall 100% Green): €115.00 / mo
  * Internet (Vodafone CableMax 1000 Wi-Fi 6): €44.99 / mo
  * GEZ Rundfunkbeitrag: €18.36 / mo (€55.08 / qtr)
  * Grundsteuer (Berlin Property Tax): €32.83 / mo (€98.50 / qtr)
  * Allianz Building Insurance: €40.00 / mo (€480.00 / yr)
* **Net Operating Income (NOI):** €1,848.15 / month (€22,177.80 / year)
* **Total Escrow Kaution Managed:** €4,932.00
* 👉 **[Open Tenant Expenses & Deposit Settlement Table](https://app.clickup.com/90122087374/v/li/1500350000002195?t=2kxv3uye-2292)**
  * **Room KE1 (Oct 31 move-out):** Deposit €1,258.00 + Expenses €38.50 - Clean fee €65.00 = **Net Return: €1,231.50** (Approved)
  * **Room KE2 (Oct 31 move-out):** Deposit €1,298.00 + Expenses €22.00 - Deductions €0.00 = **Net Return: €1,320.00** (Approved)
  * **Room KE3 (Active):** Escrow €1,238.00 | Tenant Out-of-Pocket Expense €15.00 (Pending credit)
  * **Room KE4 (Active):** Escrow €1,138.00 | Tenant Out-of-Pocket Expense €24.90 (Pending credit)


---

### 🧰 Amenities & Inventory Matrix
* **Total Itemized Amenities:** 45 Items tracked in table format
* **Table View Direct Link:** [Open Amenities Table View](https://app.clickup.com/90122087374/v/li/1500350000002184?t=2kxv3uye-2012)
* Highlights: Samsung SpaceMax™ double-door fridge, 4-mode spa standing shower, ceramic induction/oven, luggage storage room, and AVM FRITZ!Box 6660 Wi-Fi 6 mesh.

---

### 🔗 Direct 1-Click Launchpad (External Systems)
* 📊 [Looker Studio Report #1: Financial Overview & Balance Sheet](https://datastudio.google.com/u/0/reporting/49f8c6ad-5e9c-4265-931b-719c67611274/page/p_wmttjabrxd)
* 📈 [Looker Studio Report #2: Revenue Streams & Tenant Billings](https://datastudio.google.com/u/0/reporting/469f33c1-f0d1-4175-bc04-8d180a088760)
* 🎯 [Looker Studio Report #3: Occupancy & Operations KPIs](https://datastudio.google.com/u/0/reporting/012f2dfd-e933-4f6b-b34b-72ff0730a672/page/K22jF)
* 📑 [Google Sheet: Fillout Bookings & Tenant Management](https://docs.google.com/spreadsheets/d/1gcQ4nv-lu8j3Kt3mNoFx11W7jkU0Umu1Ee2C2wM6Nvw/edit)
* 🖥️ [Grafana Cloud Live Monitoring Dashboard](https://kylash.grafana.net/dashboards/f/homelyo-vgwf4nll67lpxuvevxmeog3r1vacdq6z/?orgId=1)
* 📁 [Google Drive Master Property Documents Vault](https://drive.google.com/drive/u/0/folders/1jVYbBfrm3SHqIou-Xf9iVrYaLsro5X0p)
* 🌐 [Homelyo Public Live Website](https://www.homelyo.eu/)
* 🤖 [Sara AI Assistant Cloudflare Proxy](https://assistant-proxy.kylashpriya.workers.dev/)
"""

url = "https://api.clickup.com/api/v2/task/1amhxucn0dr"
req = urllib.request.Request(url, data=json.dumps({"description": cockpit_md}).encode("utf-8"), headers=HEADERS, method="PUT")
with urllib.request.urlopen(req) as resp:
    print("Updated Cockpit Task description cleanly.")
