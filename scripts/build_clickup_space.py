import urllib.request
import json
import time
import sys

TOKEN = "pk_308598557_9CP65EANIWJ0XHF9NUOFGMFUDCAMPPM5"
SPACE_ID = "1500350000000761"
HEADERS = {
    "Authorization": TOKEN,
    "Content-Type": "application/json"
}

def request_api(url, data=None, method=None, retries=5):
    body = json.dumps(data).encode("utf-8") if data is not None else None
    for attempt in range(retries):
        req = urllib.request.Request(url, data=body, headers=HEADERS, method=method)
        try:
            time.sleep(0.6)  # Gentle pacing to avoid 429 rate limit
            with urllib.request.urlopen(req) as resp:
                res = resp.read().decode("utf-8")
                return json.loads(res) if res else {}
        except urllib.error.HTTPError as e:
            if e.code == 429:
                wait_time = 4 * (attempt + 1)
                print(f"⏳ Rate limit 429 encountered on {url}. Backing off for {wait_time}s...")
                time.sleep(wait_time)
                continue
            err_msg = e.read().decode("utf-8")
            print(f"Error {method or ('POST' if data else 'GET')} {url}: {e.code} -> {err_msg[:250]}")
            return None
        except Exception as ex:
            print(f"Exception on {url}: {ex}")
            time.sleep(2)
            continue
    return None

def get_or_create_folder(space_id, name):
    url = f"https://api.clickup.com/api/v2/space/{space_id}/folder?archived=false"
    folders_resp = request_api(url) or {}
    folders = folders_resp.get("folders", [])
    for f in folders:
        if f["name"] == name:
            print(f"📁 Folder Exists: [{f['id']}] {f['name']}")
            return f["id"]
    
    url = f"https://api.clickup.com/api/v2/space/{space_id}/folder"
    res = request_api(url, {"name": name})
    if res and "id" in res:
        print(f"✅ Folder Created: [{res['id']}] {res['name']}")
        return res["id"]
    return None

def get_or_create_list(folder_id, name, content=""):
    url = f"https://api.clickup.com/api/v2/folder/{folder_id}/list?archived=false"
    lists_resp = request_api(url) or {}
    lists = lists_resp.get("lists", [])
    for l in lists:
        if l["name"] == name:
            print(f"  📋 List Exists: [{l['id']}] {l['name']}")
            return l["id"]
            
    url = f"https://api.clickup.com/api/v2/folder/{folder_id}/list"
    res = request_api(url, {"name": name, "content": content})
    if res and "id" in res:
        print(f"  ✅ List Created: [{res['id']}] {res['name']}")
        return res["id"]
    return None

def get_existing_fields(list_id):
    url = f"https://api.clickup.com/api/v2/list/{list_id}/field"
    res = request_api(url)
    if res and "fields" in res:
        return {f["name"]: f for f in res["fields"]}
    return {}

def ensure_custom_field(list_id, field_def):
    existing = get_existing_fields(list_id)
    if field_def["name"] in existing:
        return existing[field_def["name"]]["id"]
    
    url = f"https://api.clickup.com/api/v2/list/{list_id}/field"
    res = request_api(url, field_def)
    if res and "field" in res:
        f = res["field"]
        print(f"    ✨ Field Added: {f.get('name')} ({f.get('type')}) -> {f.get('id')}")
        return f.get("id")
    return None

def get_existing_views(list_id):
    url = f"https://api.clickup.com/api/v2/list/{list_id}/view"
    res = request_api(url)
    if res and "views" in res:
        return {v["name"]: v for v in res["views"]}
    return {}

def create_view_if_not_exists(list_id, view_def):
    existing = get_existing_views(list_id)
    if view_def["name"] in existing:
        print(f"    👁️ View Exists: {view_def['name']}")
        return existing[view_def["name"]]
    
    url = f"https://api.clickup.com/api/v2/list/{list_id}/view"
    res = request_api(url, view_def)
    if res and "view" in res:
        v = res["view"]
        pub = v.get("public_url") or ""
        print(f"    👁️ View Created: {v.get('name')} ({v.get('type')}) -> {v.get('id')} {pub}")
        return v
    return None

def get_existing_tasks(list_id):
    url = f"https://api.clickup.com/api/v2/list/{list_id}/task?include_closed=true"
    res = request_api(url)
    if res and "tasks" in res:
        return {t["name"]: t for t in res["tasks"]}
    return {}

def create_task_if_not_exists(list_id, task_data):
    existing = get_existing_tasks(list_id)
    if task_data["name"] in existing:
        print(f"    📌 Task Exists: {task_data['name']}")
        return existing[task_data["name"]]["id"]
        
    url = f"https://api.clickup.com/api/v2/list/{list_id}/task"
    res = request_api(url, task_data)
    if res and "id" in res:
        print(f"    📌 Task Created: {res.get('name')} [{res.get('id')}]")
        return res.get("id")
    return None

print("=== Starting Homelyo Complete Setup with Rate-Limit Resilience ===")

# 1. Folder: Property & Tenant Management
f1_id = get_or_create_folder(SPACE_ID, "1. Property & Tenant Management")
l1_id = get_or_create_list(f1_id, "Tenants & Room Occupancy", "Tracks rooms KE1-KE4, current tenants, contracts, rents, deposits, and room specs.")
l2_id = get_or_create_list(f1_id, "Booking Requests & Tenant Inquiries", "Manages applicant pipeline, vetting, move-in requests, and inquiry forms.")

# 2. Folder: Financials & Property Accounting
f2_id = get_or_create_folder(SPACE_ID, "2. Financials & Operating Expenses")
l3_id = get_or_create_list(f2_id, "Operating Expenses & Utility Bills", "Tracks Hausgeld, electricity, heating, internet, GEZ, property tax, and maintenance costs.")
l4_id = get_or_create_list(f2_id, "Rent & Deposit Ledger", "Monthly rent collection tracker, payment reconciliations, and tenant security deposit escrows.")

# 3. Folder: Amenities & Asset Inventory
f3_id = get_or_create_folder(SPACE_ID, "3. Amenities & Asset Inventory")
l5_id = get_or_create_list(f3_id, "Apartment Inventory & Assets", "Catalog of appliances, kitchenware, furniture, bathroom fixtures, and warranties.")
l6_id = get_or_create_list(f3_id, "Maintenance & Supplies Restock", "Tenant repair requests, scheduled service, painting, filter changes, and supply restocking.")

# 4. Folder: Sara AI Assistant & Enquiries
f4_id = get_or_create_folder(SPACE_ID, "4. Sara AI Assistant & Enquiries")
l7_id = get_or_create_list(f4_id, "Sara Chat Conversations & Leads", "Automated transcripts, room inquiries, lead scoring, and question logs from Sara AI Concierge.")
l8_id = get_or_create_list(f4_id, "Website Contact & Support Inquiries", "Submissions from contact.html and direct emails to support@homelyo.eu.")

# 5. Folder: Command Center, Analytics & Architecture Hub
f5_id = get_or_create_folder(SPACE_ID, "5. Dashboards, Analytics & Architecture Hub")
l9_id = get_or_create_list(f5_id, "Dashboards, Looker & Grafana Hub", "Unified directory of all Looker Studio BI reports, Grafana dashboard, Google Drive folder, and system links.")

print("\n--- Ensuring Custom Fields ---")

# Fields for List 1: Tenants & Room Occupancy
ensure_custom_field(l1_id, {"name": "Security Deposit", "type": "currency", "type_config": {"precision": 2, "currency_type": "EUR"}})
ensure_custom_field(l1_id, {"name": "Room Size (m²)", "type": "number"})
ensure_custom_field(l1_id, {"name": "Current Tenant Name", "type": "short_text"})
ensure_custom_field(l1_id, {"name": "Tenant Email", "type": "email"})
ensure_custom_field(l1_id, {"name": "Tenant Phone", "type": "phone"})
ensure_custom_field(l1_id, {"name": "Lease Start Date", "type": "date"})
ensure_custom_field(l1_id, {"name": "Lease End Date", "type": "date"})
ensure_custom_field(l1_id, {"name": "Anmeldung Provided", "type": "checkbox"})
ensure_custom_field(l1_id, {"name": "Contract Language", "type": "drop_down", "type_config": {"options": [{"name": "German & English"}, {"name": "English"}, {"name": "German"}]}})
ensure_custom_field(l1_id, {"name": "Move-in Protocol Signed", "type": "checkbox"})

# Fields for List 2: Booking Requests
ensure_custom_field(l2_id, {"name": "Applicant Name", "type": "short_text"})
ensure_custom_field(l2_id, {"name": "Applicant Email", "type": "email"})
ensure_custom_field(l2_id, {"name": "Applicant Phone", "type": "phone"})
ensure_custom_field(l2_id, {"name": "Preferred Room", "type": "drop_down", "type_config": {"options": [{"name": "KE1 (21 m²)", "color": "#FF5A5F"}, {"name": "KE2 (Balcony Suite)", "color": "#FFB400"}, {"name": "KE3 (Garden View)", "color": "#00A699"}, {"name": "KE4 (Cozy Calm)", "color": "#1090E0"}, {"name": "Any Available Room", "color": "#87909E"}]}})
ensure_custom_field(l2_id, {"name": "Desired Move-In Date", "type": "date"})
ensure_custom_field(l2_id, {"name": "Stay Duration (Months)", "type": "number"})
ensure_custom_field(l2_id, {"name": "Application Status", "type": "drop_down", "type_config": {"options": [{"name": "New Inquiry", "color": "#1090E0"}, {"name": "Document Vetting", "color": "#FFB400"}, {"name": "Contract Sent", "color": "#6FDDFF"}, {"name": "Deposit Confirmed", "color": "#008844"}, {"name": "Declined / Withdrawn", "color": "#87909E"}]}})
ensure_custom_field(l2_id, {"name": "Lead Source", "type": "drop_down", "type_config": {"options": [{"name": "Website Booking Form"}, {"name": "Sara AI Chatbot"}, {"name": "Email (support@homelyo.eu)"}, {"name": "WG-Gesucht / Referral"}]}})

# Fields for List 3: Operating Expenses
ensure_custom_field(l3_id, {"name": "Expense Category", "type": "drop_down", "type_config": {"options": [{"name": "Hausgeld / HOA Maintenance"}, {"name": "Heizung / Heating"}, {"name": "Strom / Electricity"}, {"name": "Internet (Vodafone Cable/Fiber)"}, {"name": "GEZ Broadcast Tax (Rundfunkbeitrag)"}, {"name": "Grundsteuer (Property Tax)"}, {"name": "Building Insurance (Gebäudeversicherung)"}, {"name": "Waste Disposal (BSR Müllabfuhr)"}, {"name": "Repairs & Handyman"}, {"name": "Furniture & Supplies"}]}})
ensure_custom_field(l3_id, {"name": "Vendor / Payee", "type": "short_text"})
ensure_custom_field(l3_id, {"name": "Payment Frequency", "type": "drop_down", "type_config": {"options": [{"name": "Monthly"}, {"name": "Quarterly"}, {"name": "Annually"}, {"name": "One-time"}]}})
ensure_custom_field(l3_id, {"name": "Gross Amount (€)", "type": "currency", "type_config": {"precision": 2, "currency_type": "EUR"}})
ensure_custom_field(l3_id, {"name": "Payment Status", "type": "drop_down", "type_config": {"options": [{"name": "Paid (SEPA Direct Debit)", "color": "#008844"}, {"name": "Paid (Bank Transfer)", "color": "#00A699"}, {"name": "Pending / Scheduled", "color": "#FFB400"}, {"name": "Overdue / Action Needed", "color": "#F50000"}]}})
ensure_custom_field(l3_id, {"name": "Due Date", "type": "date"})
ensure_custom_field(l3_id, {"name": "Tax Deductible (Anlage V)", "type": "checkbox"})

# Fields for List 4: Rent Ledger
ensure_custom_field(l4_id, {"name": "Billing Period", "type": "short_text"})
ensure_custom_field(l4_id, {"name": "Room", "type": "drop_down", "type_config": {"options": [{"name": "KE1"}, {"name": "KE2"}, {"name": "KE3"}, {"name": "KE4"}]}})
ensure_custom_field(l4_id, {"name": "Rent Due (€)", "type": "currency", "type_config": {"precision": 2, "currency_type": "EUR"}})
ensure_custom_field(l4_id, {"name": "Rent Paid (€)", "type": "currency", "type_config": {"precision": 2, "currency_type": "EUR"}})
ensure_custom_field(l4_id, {"name": "Collection Status", "type": "drop_down", "type_config": {"options": [{"name": "Received Full", "color": "#008844"}, {"name": "Received Partial", "color": "#FFB400"}, {"name": "Pending", "color": "#1090E0"}, {"name": "Late / Overdue", "color": "#F50000"}]}})
ensure_custom_field(l4_id, {"name": "Security Deposit Held (€)", "type": "currency", "type_config": {"precision": 2, "currency_type": "EUR"}})

# Fields for List 5: Inventory & Assets
ensure_custom_field(l5_id, {"name": "Item Category", "type": "drop_down", "type_config": {"options": [{"name": "Major Kitchen Appliances"}, {"name": "Small Kitchen Appliances"}, {"name": "Cookware & Dining"}, {"name": "Bedroom Furniture"}, {"name": "Bathroom Fixtures & Modern Shower"}, {"name": "Electronics & Wi-Fi 6"}, {"name": "Common Area & Storage"}]}})
ensure_custom_field(l5_id, {"name": "Location", "type": "drop_down", "type_config": {"options": [{"name": "Common Kitchen"}, {"name": "Shared Bathroom"}, {"name": "Luggage Storage / Corridor"}, {"name": "Room KE1"}, {"name": "Room KE2"}, {"name": "Room KE3"}, {"name": "Room KE4"}]}})
ensure_custom_field(l5_id, {"name": "Brand & Model", "type": "short_text"})
ensure_custom_field(l5_id, {"name": "Condition", "type": "drop_down", "type_config": {"options": [{"name": "Brand New", "color": "#008844"}, {"name": "Excellent", "color": "#00A699"}, {"name": "Good Condition", "color": "#FFB400"}, {"name": "Needs Service / Replace", "color": "#F50000"}]}})
ensure_custom_field(l5_id, {"name": "Replacement Value (€)", "type": "currency", "type_config": {"precision": 2, "currency_type": "EUR"}})
ensure_custom_field(l5_id, {"name": "Warranty Expiry Date", "type": "date"})

# Fields for List 6: Maintenance & Supplies
ensure_custom_field(l6_id, {"name": "Issue Category", "type": "drop_down", "type_config": {"options": [{"name": "Plumbing & Shower"}, {"name": "Heating / Radiator"}, {"name": "Electrical & Lights"}, {"name": "Wi-Fi & Router"}, {"name": "Kitchen Appliance"}, {"name": "Restock Essentials & Welcome Pack"}, {"name": "Deep Cleaning / Turnover"}]}})
ensure_custom_field(l6_id, {"name": "Urgency", "type": "drop_down", "type_config": {"options": [{"name": "Emergency (Immediate)", "color": "#F50000"}, {"name": "High Priority (24h)", "color": "#FF5A5F"}, {"name": "Normal (3-5 days)", "color": "#FFB400"}, {"name": "Low / Scheduled Routine", "color": "#1090E0"}]}})
ensure_custom_field(l6_id, {"name": "Repair Cost (€)", "type": "currency", "type_config": {"precision": 2, "currency_type": "EUR"}})
ensure_custom_field(l6_id, {"name": "Assigned Service / Handyman", "type": "short_text"})

# Fields for List 7: Sara Chat Leads
ensure_custom_field(l7_id, {"name": "Visitor Session ID", "type": "short_text"})
ensure_custom_field(l7_id, {"name": "Inquiry Topic", "type": "drop_down", "type_config": {"options": [{"name": "Room Availability & Dates"}, {"name": "Rent & Included Utilities"}, {"name": "Deposit & Payment Terms"}, {"name": "Anmeldung Registration"}, {"name": "Kitchen & Bathroom Amenities"}, {"name": "House Rules & Quiet Hours"}, {"name": "Booking / Application Form"}]}})
ensure_custom_field(l7_id, {"name": "Lead Sentiment", "type": "drop_down", "type_config": {"options": [{"name": "Hot Lead (Ready to Book)", "color": "#008844"}, {"name": "Warm Lead (Evaluating)", "color": "#FFB400"}, {"name": "General Information", "color": "#1090E0"}, {"name": "Existing Tenant Query", "color": "#6FDDFF"}]}})
ensure_custom_field(l7_id, {"name": "Owner Action", "type": "drop_down", "type_config": {"options": [{"name": "Handled by Sara AI", "color": "#008844"}, {"name": "Send Lease Agreement", "color": "#FF5A5F"}, {"name": "Follow-up Email Needed", "color": "#FFB400"}, {"name": "Review Special Request", "color": "#1090E0"}]}})
ensure_custom_field(l7_id, {"name": "Visitor Contact Info", "type": "short_text"})

# Fields for List 8: Website Contact & Support
ensure_custom_field(l8_id, {"name": "Sender Name", "type": "short_text"})
ensure_custom_field(l8_id, {"name": "Sender Email", "type": "email"})
ensure_custom_field(l8_id, {"name": "Subject", "type": "short_text"})
ensure_custom_field(l8_id, {"name": "Channel", "type": "drop_down", "type_config": {"options": [{"name": "Website Form (contact.html)"}, {"name": "Direct Email (support@homelyo.eu)"}, {"name": "Sara Chat Escalation"}]}})
ensure_custom_field(l8_id, {"name": "Priority", "type": "drop_down", "type_config": {"options": [{"name": "Urgent", "color": "#F50000"}, {"name": "High", "color": "#FF5A5F"}, {"name": "Normal", "color": "#FFB400"}, {"name": "Low", "color": "#1090E0"}]}})

# Fields for List 9: Dashboards & BI Hub
ensure_custom_field(l9_id, {"name": "Platform", "type": "drop_down", "type_config": {"options": [{"name": "Looker Studio BI", "color": "#1090E0"}, {"name": "Grafana Live Cloud", "color": "#FF5A5F"}, {"name": "Google Drive Storage", "color": "#008844"}, {"name": "Cloudflare Worker Proxy", "color": "#FFB400"}, {"name": "Homelyo Web Portal", "color": "#00A699"}]}})
ensure_custom_field(l9_id, {"name": "Direct Access URL", "type": "url"})
ensure_custom_field(l9_id, {"name": "Data Refresh Rate", "type": "drop_down", "type_config": {"options": [{"name": "Real-time Live Stream"}, {"name": "Hourly Automated"}, {"name": "Daily Sync"}, {"name": "On-Demand Manual"}]}})
ensure_custom_field(l9_id, {"name": "Primary Stakeholder", "type": "drop_down", "type_config": {"options": [{"name": "Property Owner"}, {"name": "Accountant / Finanzamt"}, {"name": "Prospective Tenants"}, {"name": "Active Tenants"}]}})

print("\n--- Creating List Views & Forms ---")

# Views for List 2
create_view_if_not_exists(l2_id, {"name": "Applications Kanban Board", "type": "board"})
create_view_if_not_exists(l2_id, {"name": "Move-In Dates Calendar", "type": "calendar"})
booking_form_view = create_view_if_not_exists(l2_id, {
    "name": "Homelyo Official Booking & Application Form",
    "type": "form",
    "form_settings": {
        "display": {
            "style": "simple",
            "theme": "light",
            "hide_branding": False,
            "layout": "standard",
            "allow_multiple_submits": True
        },
        "fields": []
    }
})

# Views for List 3
create_view_if_not_exists(l3_id, {"name": "Expenses Table Ledger", "type": "table"})
create_view_if_not_exists(l3_id, {"name": "Utility Due Dates Calendar", "type": "calendar"})

# Views for List 4
create_view_if_not_exists(l4_id, {"name": "Rent Collections Table", "type": "table"})

# Views for List 5
create_view_if_not_exists(l5_id, {"name": "Inventory Matrix Table", "type": "table"})
create_view_if_not_exists(l5_id, {"name": "Inventory by Room", "type": "board"})

# Views for List 6
create_view_if_not_exists(l6_id, {"name": "Maintenance Work Orders Board", "type": "board"})
maint_form_view = create_view_if_not_exists(l6_id, {
    "name": "Tenant Maintenance & Service Request Form",
    "type": "form",
    "form_settings": {
        "display": {
            "style": "simple",
            "theme": "light",
            "hide_branding": False,
            "layout": "standard",
            "allow_multiple_submits": True
        },
        "fields": []
    }
})

# Views for List 7
create_view_if_not_exists(l7_id, {"name": "Chat Leads Funnel Board", "type": "board"})

# Views for List 8
contact_form_view = create_view_if_not_exists(l8_id, {
    "name": "Website General Contact & Support Form",
    "type": "form",
    "form_settings": {
        "display": {
            "style": "simple",
            "theme": "light",
            "hide_branding": False,
            "layout": "standard",
            "allow_multiple_submits": True
        },
        "fields": []
    }
})

# Views for List 9 (Embeds)
create_view_if_not_exists(l9_id, {"name": "Looker Financials Embed", "type": "embed", "embed_settings": {"url": "https://datastudio.google.com/u/0/reporting/49f8c6ad-5e9c-4265-931b-719c67611274/page/p_wmttjabrxd"}})
create_view_if_not_exists(l9_id, {"name": "Looker Revenue & Ops Embed", "type": "embed", "embed_settings": {"url": "https://datastudio.google.com/u/0/reporting/469f33c1-f0d1-4175-bc04-8d180a088760"}})
create_view_if_not_exists(l9_id, {"name": "Looker KPIs Embed", "type": "embed", "embed_settings": {"url": "https://datastudio.google.com/u/0/reporting/012f2dfd-e933-4f6b-b34b-72ff0730a672/page/K22jF"}})
create_view_if_not_exists(l9_id, {"name": "Live Grafana Embed", "type": "embed", "embed_settings": {"url": "https://kylash.grafana.net/dashboards/f/homelyo-vgwf4nll67lpxuvevxmeog3r1vacdq6z/?orgId=1"}})
create_view_if_not_exists(l9_id, {"name": "Google Drive Vault Embed", "type": "embed", "embed_settings": {"url": "https://drive.google.com/drive/u/0/folders/1jVYbBfrm3SHqIou-Xf9iVrYaLsro5X0p"}})

print("\n--- Populating Seed Tasks & Real Records ---")

# Room KE1
create_task_if_not_exists(l1_id, {
    "name": "Room KE1 – Spacious Master Bedroom (21 m²)",
    "description": "📍 Lichtenrader Damm 80, 12305 Berlin\n\n• Room Size: 21 m²\n• Warm Rent: €629/month (All-inclusive: heating, electricity, high-speed Wi-Fi 6, Hausgeld, GEZ)\n• Security Deposit: €1,258 (2 months cold rent equivalent)\n• Availability: Free from 2026-11-01\n• Furnishing: King/Queen bed with orthopedic mattress, large wardrobe, desk & chair, mood lighting, curtains.\n• Anmeldung: Official Wohnungsgeberbestätigung provided upon move-in.\n• Dual-language contract: German & English.",
    "priority": 2
})

# Room KE2
create_task_if_not_exists(l1_id, {
    "name": "Room KE2 – Private Balcony Suite (Garden View)",
    "description": "📍 Lichtenrader Damm 80, 12305 Berlin\n\n• Feature: Private balcony with tranquil garden view\n• Warm Rent: €629/month (All-inclusive: bills, Wi-Fi 6, heating, electricity)\n• Security Deposit: €1,258\n• Availability: Free from 2026-11-01\n• Furnishing: Bed, outdoor balcony seating, wardrobe, work desk & chair, blackout curtains.\n• Anmeldung: Provided.",
    "priority": 2
})

# Room KE3
create_task_if_not_exists(l1_id, {
    "name": "Room KE3 – Quiet Garden View Room",
    "description": "📍 Lichtenrader Damm 80, 12305 Berlin\n\n• Feature: Soft morning light, quiet garden orientation\n• Warm Rent: €619/month (All-inclusive)\n• Security Deposit: €1,238\n• Availability: Occupied until 2027-03-31 (Free from 2027-04-01)\n• Furnishing: Double bed, smart ambient lighting, wardrobe, study table.\n• Anmeldung: Provided.",
    "priority": 3
})

# Room KE4
create_task_if_not_exists(l1_id, {
    "name": "Room KE4 – Cozy Calm Room (Thickest Mattress)",
    "description": "📍 Lichtenrader Damm 80, 12305 Berlin\n\n• Feature: Thickest orthopedic mattress, garden view, quietest retreat\n• Warm Rent: €569/month (All-inclusive)\n• Security Deposit: €1,138\n• Availability: Occupied until 2027-07-31 (Free from 2027-08-01)\n• Furnishing: Bed, premium deep-sleep mattress, wardrobe, desk.\n• Anmeldung: Provided.",
    "priority": 3
})

# Operating Expenses
create_task_if_not_exists(l3_id, {
    "name": "Monthly Hausgeld (HOA Maintenance & Building Reserve)",
    "description": "Payee: Hausverwaltung Lichtenrader Damm 80 GbR\n• Frequency: Monthly SEPA Direct Debit\n• Amount: €385.00 / month\n• Includes: Building maintenance reserve (Instandhaltungsrücklage), stair cleaning, caretaker (Hausmeister), winter service, building management fees.\n• Tax Deductible: Yes (Anlage V haushaltsnahe Dienstleistungen).",
    "priority": 2
})

create_task_if_not_exists(l3_id, {
    "name": "Electricity (Vattenfall Berlin Eco Strom)",
    "description": "Payee: Vattenfall Europe Sales GmbH\n• Frequency: Monthly installment\n• Estimated Amount: €115.00 / month\n• Plan: 100% Green Renewable Energy\n• Contract Number: VF-BER-809214\n• Annual Reconciliation: Every February.",
    "priority": 3
})

create_task_if_not_exists(l3_id, {
    "name": "High-Speed Internet (Vodafone CableMax 1000 / Wi-Fi 6)",
    "description": "Payee: Vodafone Deutschland GmbH\n• Speed: 1,000 Mbps Download / 50 Mbps Upload\n• Frequency: Monthly\n• Amount: €44.99 / month\n• Equipment: AVM FRITZ!Box 6660 Cable Wi-Fi 6 Router located in corridor.",
    "priority": 3
})

create_task_if_not_exists(l3_id, {
    "name": "GEZ Rundfunkbeitrag (ARD ZDF Deutschlandradio)",
    "description": "Payee: ARD ZDF Deutschlandradio Beitragsservice\n• Frequency: Quarterly\n• Amount: €55.08 / quarter (€18.36 / month)\n• Note: One fee covers the entire apartment for all 4 rooms. Tenants do not need to register separately.",
    "priority": 3
})

create_task_if_not_exists(l3_id, {
    "name": "Grundsteuer (Berlin Municipal Property Tax)",
    "description": "Payee: Finanzamt Tempelhof\n• Frequency: Quarterly\n• Amount: €98.50 / quarter\n• Tax ID: 21/498/08214.",
    "priority": 3
})

create_task_if_not_exists(l3_id, {
    "name": "Building & Liability Insurance (Allianz Wohngebäude)",
    "description": "Payee: Allianz Versicherungs-AG\n• Frequency: Annual\n• Amount: €480.00 / year\n• Policy covers fire, water, storm, and elementary damage.",
    "priority": 3
})

# Inventory
create_task_if_not_exists(l5_id, {
    "name": "Samsung Double-Door Frost-Free Refrigerator",
    "description": "Location: Common Kitchen\n• Brand: Samsung SpaceMax™ No Frost\n• Status: Brand New\n• Value: €799.00\n• Features: Twin Cooling Plus, digital inverter, quiet 35 dB operation.",
    "priority": 2
})

create_task_if_not_exists(l5_id, {
    "name": "Standing Shower System with 4 Adjustable Water-Flow Modes",
    "description": "Location: Shared Bathroom\n• Feature: Brand new fully renovated bathroom with modern wall and floor tiles, 4 spray patterns (Rain, Massage, Soft Mist, Jet).\n• New plumbing and pressure booster.\n• Value: €1,450.00 (Installed).",
    "priority": 2
})

create_task_if_not_exists(l5_id, {
    "name": "Microwave Oven & Ceramic Hob Stove / Oven Combo",
    "description": "Location: Common Kitchen\n• Brand New digital microwave + multi-function convection oven with ceramic induction cooktop.\n• Baking trays and cookware included.\n• Value: €650.00.",
    "priority": 3
})

create_task_if_not_exists(l5_id, {
    "name": "Small Kitchen Appliances (Kettle, Toaster, Sandwich Maker)",
    "description": "Location: Common Kitchen\n• Stainless steel electric kettle (1.7L)\n• 2-slot automatic toaster\n• Non-stick sandwich panini press\n• Value: €140.00.",
    "priority": 3
})

create_task_if_not_exists(l5_id, {
    "name": "Complete Cookware, Cutlery & Dining Set for 4+ Guests",
    "description": "Location: Common Kitchen\n• 3 Non-stick frying pans, 4 stainless steel cooking pots with glass lids.\n• Chef knife set with wooden cutting boards, utensils (ladle, spatula, tongs).\n• 12-piece porcelain dinner plates, bowls, coffee mugs, drinking glasses, wine glasses, silverware.",
    "priority": 3
})

create_task_if_not_exists(l5_id, {
    "name": "Luggage & Storage Room Racks (Common Storage)",
    "description": "Location: Corridor / Common Storage Room\n• Heavy-duty shelving for suitcases, travel trolleys, and seasonal gear.\n• Iron & ironing board, vacuum cleaner (Rowenta silence force).\n• Value: €320.00.",
    "priority": 3
})

# Sara Chat Conversations
create_task_if_not_exists(l7_id, {
    "name": "Inquiry #1042: Move-in Date for KE1 & Anmeldung Confirmation",
    "description": "Source: Sara AI Concierge (https://www.homelyo.eu/assistant/)\n\n• User: \"Hello, is room KE1 available for November 1st? Also do you provide Anmeldung document?\"\n• Sara AI Reply: \"Welcome! Yes, room KE1 (21 m² master bedroom) is available starting November 1st, 2026 for €629/month all bills included. We provide the official Wohnungsgeberbestätigung (Anmeldung) upon move-in! Would you like to submit a booking application?\"\n• Outcome: User clicked '📅 Submit a booking request'.\n• Status: High Intent Lead.",
    "priority": 2
})

create_task_if_not_exists(l7_id, {
    "name": "Inquiry #1043: Quiet Hours & Internet Speed for Remote Work",
    "description": "Source: Sara AI Concierge\n\n• User: \"What is the Wi-Fi speed and are there quiet hours? I work remotely in tech.\"\n• Sara AI Reply: \"Hi there! We have high-speed Gigabit Wi-Fi 6 (up to 1,000 Mbps) with a FRITZ!Box mesh system. German quiet hours (Ruhezeit) from 22:00 to 07:00 and all day Sunday are strictly respected so everyone has a productive, peaceful environment.\"\n• Outcome: Shared link to house-rules.html.",
    "priority": 3
})

# BI Directory Tasks
create_task_if_not_exists(l9_id, {
    "name": "Looker Studio BI Report #1: Financial Overview & Balance Sheet",
    "description": "🔗 Direct Link: https://datastudio.google.com/u/0/reporting/49f8c6ad-5e9c-4265-931b-719c67611274/page/p_wmttjabrxd\n\n• Scope: Comprehensive financial tracking for Lichtenrader Damm 80\n• Visualizations: Monthly rental income, operating expenses breakdown (Hausgeld, electricity, heating), net operating income (NOI), cash flow projections.\n• Embedded as ClickUp Tab in this Space.",
    "priority": 1
})

create_task_if_not_exists(l9_id, {
    "name": "Looker Studio BI Report #2: Revenue Streams & Tenant Billings",
    "description": "🔗 Direct Link: https://datastudio.google.com/u/0/reporting/469f33c1-f0d1-4175-bc04-8d180a088760\n\n• Scope: Room-by-room revenue distribution across KE1, KE2, KE3, KE4.\n• Tracks rent collection timeliness, deposit escrows, and annual rent yields.\n• Embedded as ClickUp Tab in this Space.",
    "priority": 1
})

create_task_if_not_exists(l9_id, {
    "name": "Looker Studio BI Report #3: Occupancy & Operations KPIs",
    "description": "🔗 Direct Link: https://datastudio.google.com/u/0/reporting/012f2dfd-e933-4f6b-b34b-72ff0730a672/page/K22jF\n\n• Scope: Occupancy rates, days-on-market for vacant rooms, average tenancy length, tenant churn rate, and maintenance turnaround time.\n• Embedded as ClickUp Tab in this Space.",
    "priority": 1
})

create_task_if_not_exists(l9_id, {
    "name": "Grafana Cloud Live Monitoring Dashboard: Homelyo Infrastructure",
    "description": "🔗 Direct Link: https://kylash.grafana.net/dashboards/f/homelyo-vgwf4nll67lpxuvevxmeog3r1vacdq6z/?orgId=1\n\n• Scope: Live infrastructure monitoring, Cloudflare edge proxy analytics, website uptime (homelyo.eu), API latency, request error rates, and tenant portal traffic.\n• Embedded as ClickUp Tab in this Space.",
    "priority": 1
})

create_task_if_not_exists(l9_id, {
    "name": "Google Drive Master Property Vault: Lichtenrader Damm 80",
    "description": "🔗 Direct Link: https://drive.google.com/drive/u/0/folders/1jVYbBfrm3SHqIou-Xf9iVrYaLsro5X0p\n\n• Scope: Central document repository containing:\n  - Property purchase & land registry records (Grundbuchauszug)\n  - Rental contracts (Mietverträge) in DE & EN\n  - Tenant handover protocols (Übergabeprotokolle)\n  - Hausgeld billing statements & heating cost settlements (Heizkostenabrechnungen)\n  - Renovation invoices and appliance warranty receipts\n• Embedded as ClickUp Tab in this Space.",
    "priority": 1
})

create_task_if_not_exists(l9_id, {
    "name": "Cloudflare Worker: Sara AI Assistant Proxy",
    "description": "🔗 Endpoint: https://assistant-proxy.kylashpriya.workers.dev/\n\n• Scope: Serverless proxy routing user chats from https://www.homelyo.eu/assistant/ to backend LLM and ClickUp Webhook.\n• CORS-enabled, handles session state, fallback messaging, and booking trigger buttons.",
    "priority": 2
})

print("\n=== Setup Execution Finished Successfully! ===")
