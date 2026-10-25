import urllib.request
import json

TOKEN = "pk_308598557_9CP65EANIWJ0XHF9NUOFGMFUDCAMPPM5"
LIST_ID = "1500350000002188" # Dashboards, Looker & Grafana Hub
HEADERS = {"Authorization": TOKEN, "Content-Type": "application/json"}

cockpit_md = """# 🚀 Homelyo Executive Cockpit — Quick Glance

### 📍 Property Overview
* **Address:** Lichtenrader Damm 80, 12305 Berlin-Lichtenrade
* **Total Property Size:** 96 m² | 4 Private Bedrooms + Shared Kitchen + Renovated Luxury Shower + Dedicated Common Luggage Storage Room
* **Target Tenancy:** Professionals, Master Students, and Expats in Berlin

---

### 🛏️ Live Room Occupancy & Free Dates
| Room | Feature | Size | Warm Rent (All-inclusive) | Current Status | Available From |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **KE1** | Master Bedroom | 21 m² | **€629 / mo** | 🟢 **Available** | **Nov 1, 2026** |
| **KE2** | Private Balcony Suite | ~18 m² | **€649 / mo** | 🟢 **Available** | **Nov 1, 2026** |
| **KE3** | Quiet Garden View | ~16 m² | **€619 / mo** | 🟡 Occupied | **Apr 1, 2027** |
| **KE4** | Cozy Calm (Thick Mattress) | ~14 m² | **€569 / mo** | 🟡 Occupied | **Aug 1, 2027** |

---

### 💰 Financial Rollup & Cash Flow
* **Gross Potential Monthly Revenue (100% Occupancy):** `€2,466.00 / month` (`€29,592.00 / year`)
* **Total Operating Expenses (All Utilities + Hausgeld):** `€617.85 / month`
  * Hausgeld (HOA, Reserve, Maintenance): €385.00 / mo
  * Electricity (Vattenfall 100% Green): €115.00 / mo
  * Internet (Vodafone CableMax 1000 Wi-Fi 6): €44.99 / mo
  * GEZ Rundfunkbeitrag: €18.36 / mo (€55.08 / qtr)
  * Grundsteuer (Berlin Property Tax): €32.83 / mo (€98.50 / qtr)
  * Allianz Building Insurance: €40.00 / mo (€480.00 / yr)
* **Net Operating Income (NOI):** **`€1,848.15 / month`** (`€22,177.80 / year`)
* **Tenant Security Deposits Held in Escrow:** **`€4,932.00`**

---

### 🧰 Amenities & Inventory Matrix
* **Total Itemized Amenities:** **45 Items** tracked in table format
* **Table View Direct Link:** [Open Amenities Table View](https://app.clickup.com/90122087374/v/li/1500350000002184?t=2kxv3uye-2012)
* Highlights: Samsung SpaceMax™ double-door fridge, 4-mode spa standing shower, ceramic induction/oven, luggage storage room, and AVM FRITZ!Box 6660 Wi-Fi 6 mesh.

---

### 🔗 Direct 1-Click Launchpad (External Systems)
* 📊 [Looker Studio Report #1: Financial Overview & Balance Sheet](https://datastudio.google.com/u/0/reporting/49f8c6ad-5e9c-4265-931b-719c67611274/page/p_wmttjabrxd)
* 📈 [Looker Studio Report #2: Revenue Streams & Tenant Billings](https://datastudio.google.com/u/0/reporting/469f33c1-f0d1-4175-bc04-8d180a088760)
* 🎯 [Looker Studio Report #3: Occupancy & Operations KPIs](https://datastudio.google.com/u/0/reporting/012f2dfd-e933-4f6b-b34b-72ff0730a672/page/K22jF)
* 🖥️ [Grafana Cloud Live Monitoring Dashboard](https://kylash.grafana.net/dashboards/f/homelyo-vgwf4nll67lpxuvevxmeog3r1vacdq6z/?orgId=1)
* 📁 [Google Drive Master Property Documents Vault](https://drive.google.com/drive/u/0/folders/1jVYbBfrm3SHqIou-Xf9iVrYaLsro5X0p)
* 🌐 [Homelyo Public Live Website](https://www.homelyo.eu/)
* 🤖 [Sara AI Assistant Cloudflare Proxy](https://assistant-proxy.kylashpriya.workers.dev/)
"""

payload = {
    "name": "⭐ Homelyo Quick-Glance Executive Cockpit & Master Directory",
    "description": cockpit_md,
    "priority": 1
}

url = f"https://api.clickup.com/api/v2/list/{LIST_ID}/task"
req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=HEADERS)
with urllib.request.urlopen(req) as resp:
    data = json.loads(resp.read().decode())
    print("Created Cockpit Task:", data.get("id"), data.get("name"))
