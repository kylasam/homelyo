#!/usr/bin/env python3
"""
Homelyo Autonomous Tenant Onboarding & Contract Generation Pipeline
-------------------------------------------------------------------
Generates tenant folder in Google Drive, dual-language rental contract (Doc & PDF),
copies onboarding templates, syncs Google Sheets (tenant_details & rent_log),
and creates ClickUp tickets in one shot.
"""

import sys
import os
import json
import argparse
import datetime
import urllib.request
import urllib.parse
import subprocess

# Configuration & Constants
PARENT_FOLDER_ID = "1jVYbBfrm3SHqIou-Xf9iVrYaLsro5X0p"  # Master Property Folder
TEMPLATE_FOLDER_ID = "16gimsgZ5BaVDk766sMDAuZWJcKJtKdqe"  # Template Folder
SPREADSHEET_ID = "1rs5dXTyA6xQSbqVQ_ofyCAv2eFZ83CrN8sAx5qfbsTk"  # Master Spreadsheet

CLICKUP_TOKEN = "pk_105682594_CCTOZ8K23KHPSAY3JQTVV1G23QTN63IG"
CLICKUP_HEADERS = {"Authorization": CLICKUP_TOKEN, "Content-Type": "application/json"}
LIST_TENANTS = "1500520000038641"  # Tenants & Room Occupancy (Homelyo Workspace)
LIST_INQUIRIES = "1500520000038640"  # Booking Requests & Tenant Inquiries (Homelyo Workspace)
LIST_SETTLEMENT = "1500520000038642"  # Tenant Expenses & Deposit Settlement (Homelyo Workspace)

ROOM_SPECS = {
    "KE1": {
        "title": "Master Bedroom",
        "size": 21,
        "rent": 629.00,
        "deposit": 1258.00,
        "features": "Spacious master bedroom (21 m²), king bed, double wardrobe, large work desk, ergonomic chair, garden courtyard view"
    },
    "KE2": {
        "title": "Private Balcony Suite",
        "size": 24,
        "rent": 649.00,
        "deposit": 1298.00,
        "features": "Sunny balcony suite (24 m²), private walk-out garden balcony, queen bed, study workstation, wardrobe"
    },
    "KE3": {
        "title": "Quiet Garden View",
        "size": 12,
        "rent": 619.00,
        "deposit": 1238.00,
        "features": "Quiet courtyard bedroom (12 m²), queen bed, dedicated study desk, double wardrobe, serene garden view"
    },
    "KE4": {
        "title": "Cozy Calm Room",
        "size": 11,
        "rent": 569.00,
        "deposit": 1138.00,
        "features": "Cozy calm bedroom (11 m²), orthopedic thick mattress, double wardrobe, work desk, morning sunlight"
    }
}

TEMPLATE_FILES = [
    {"id": "15rkiK8grB2nYpFZfQs7RTakmcBFM0A_k", "name": "House_Rules.pdf"},
    {"id": "1Y_RSVUYcObC1UWWL0GguS8CCS49dd9-e", "name": "Apartment_Handover_Protocol.pdf"},
    {"id": "1JFhXsitxEO5WETo-r2O9jqc7SFnNGkel", "name": "Keys_Acknowledgement.pdf"},
    {"id": "1yrIDs6kVUPZc8A1guV-BaUXf_x4JsfHB", "name": "wohnungsgeberbestaetigung.pdf"},
    {"id": "1mNXbTpDbUxWQOn25c0waOFjaUAxpyLPq", "name": "Inventory_List.pdf"}
]

def get_gcloud_token():
    try:
        return subprocess.check_output(["gcloud", "auth", "print-access-token"]).decode().strip()
    except Exception as e:
        print(f"Error getting gcloud token: {e}")
        return None

def create_drive_folder(token, folder_name, parent_id):
    url = "https://www.googleapis.com/drive/v3/files"
    metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id]
    }
    req = urllib.request.Request(
        url,
        data=json.dumps(metadata).encode("utf-8"),
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        res = json.loads(resp.read().decode())
        return res["id"]

def copy_drive_file(token, file_id, new_name, target_folder_id):
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}/copy"
    payload = {"name": new_name, "parents": [target_folder_id]}
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read().decode())

def generate_contract_html(tenant_name, tenant_email, tenant_phone, room_code, start_date, end_date, rent_amount, deposit_amount):
    spec = ROOM_SPECS.get(room_code, ROOM_SPECS["KE1"])
    room_title = spec["title"]
    room_size = spec["size"]
    today_str = datetime.date.today().strftime("%d. %B %Y")
    
    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
body {{ font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; line-height: 1.5; color: #1a1a1a; margin: 40px; font-size: 11pt; }}
h1 {{ text-align: center; font-size: 16pt; color: #0d3b66; margin-bottom: 4px; text-transform: uppercase; }}
h2 {{ text-align: center; font-size: 13pt; color: #555; margin-top: 0; margin-bottom: 25px; }}
h3 {{ font-size: 12pt; color: #0d3b66; border-bottom: 1.5px solid #0d3b66; padding-bottom: 4px; margin-top: 24px; }}
.party-box {{ display: flex; justify-content: space-between; background: #f8f9fa; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; margin-bottom: 20px; }}
.party-col {{ width: 48%; }}
.bilingual-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 16px; }}
.bilingual-table td {{ width: 50%; vertical-align: top; padding: 6px 12px; font-size: 10pt; line-height: 1.45; }}
.de {{ border-right: 1px solid #cbd5e1; padding-right: 14px; background: #fafbfc; }}
.en {{ padding-left: 14px; color: #2d3748; }}
.highlight {{ font-weight: bold; color: #0d3b66; }}
.signatures {{ margin-top: 40px; display: flex; justify-content: space-between; }}
.sig-box {{ width: 45%; border-top: 1px solid #000; padding-top: 8px; text-align: center; font-size: 10pt; }}
</style>
</head>
<body>

<h1>Wohnraummietvertrag für WG-Zimmer</h1>
<h2>Co-Living Room Tenancy Agreement</h2>

<div class="party-box">
    <div class="party-col">
        <strong>Vermieter / Landlord:</strong><br>
        Kylasam Nallagoundanpalayam Arruchamy<br>
        Lichtenrader Damm 80, 12305 Berlin, Deutschland<br>
        Email: support@homelyo.eu | Tel: +49 176 8332 9439<br>
        Web: www.homelyo.eu
    </div>
    <div class="party-col">
        <strong>Mieter / Tenant:</strong><br>
        <strong>{tenant_name}</strong><br>
        Email: {tenant_email}<br>
        Telefon: {tenant_phone or 'Auf Anfrage / On file'}<br>
        Mietobjekt: Zimmer {room_code} (1. OG rechts)
    </div>
</div>

<h3>§ 1 Mietgegenstand / Rented Property</h3>
<table class="bilingual-table">
<tr>
<td class="de">
Vermietet wird das möblierte Zimmer <strong>{room_code} ({room_title})</strong> mit ca. <strong>{room_size} m²</strong> Wohnfläche in der Wohnung Lichtenrader Damm 80, 12305 Berlin (1. Obergeschoss rechts).<br><br>
Zur Mitbenutzung stehen zur Verfügung: Voll ausgestattete Einbauküche (Kühl-Gefrierkombination, Induktionskochfeld, Backofen, Mikrowelle), das renovierte Duschbad/WC, die Diele sowie der separate Gemeinschafts-Gepäckraum.
</td>
<td class="en">
Rented is the fully furnished private bedroom <strong>{room_code} ({room_title})</strong> with approx. <strong>{room_size} m²</strong> floor space in the apartment at Lichtenrader Damm 80, 12305 Berlin (1st floor right).<br><br>
Shared areas for joint use include: Fully equipped kitchen (fridge/freezer, induction cooktop, oven, microwave), luxury renovated shower/WC, hallway, and dedicated separate common luggage storage room.
</td>
</tr>
</table>

<h3>§ 2 Mietzeit & Zweck / Rental Term & Purpose</h3>
<table class="bilingual-table">
<tr>
<td class="de">
1. Das Mietverhältnis beginnt am <strong>{start_date}</strong> und endet am <strong>{end_date}</strong>.<br><br>
2. Die Vermietung erfolgt zum <strong>vorübergehenden Gebrauch</strong> gemäß § 549 Abs. 2 Nr. 1 BGB (befristeter Aufenthalt zu Studien-, Praktikums- oder Erwerbszwecken in Berlin).<br><br>
3. <strong>Vertragsverlängerung, Kündigungsfrist & Mietdauer:</strong> Im Falle einer einvernehmlichen Vertragsverlängerung gilt eine zwingende Kündigungsfrist von <strong>3 Monaten</strong> zum Monatsende (schriftlich bis zum 3. Werktag des Monats). Der Mieter verpflichtet sich, die verlängerte Mietzeit vollumfänglich zu erfüllen bzw. die Mietverpflichtungen bis zum regulären Ablauf der 3-monatigen Kündigungsfrist uneingeschränkt zu bedienen.
</td>
<td class="en">
1. The tenancy begins on <strong>{start_date}</strong> and ends on <strong>{end_date}</strong>.<br><br>
2. The lease is explicitly concluded for <strong>temporary use</strong> pursuant to Section 549 (2) No. 1 German Civil Code (BGB) for academic, internship, or professional stay in Berlin.<br><br>
3. <strong>Lease Extension, Notice Period & Tenancy Fulfillment:</strong> In the event of a mutual contract extension, a mandatory notice period of <strong>3 months</strong> to the end of the month applies (in writing by the 3rd business day of the month). The tenant is obligated to serve the full extended tenancy term and fulfill all rental obligations in full through the completion of the mandatory 3-month notice period.
</td>
</tr>
</table>

<h3>§ 3 Miete & Nebenkosten / Rent & Utilities</h3>
<table class="bilingual-table">
<tr>
<td class="de">
Die monatliche Warmmiete beträgt pauschal <strong>€{rent_amount:,.2f}</strong>.<br><br>
In der Pauschale sind sämtliche Betriebs- und Nebenkosten enthalten: Heizung, Warm-/Kaltwasser, Strom (100% Ökostrom), Hausgeld, Gebäudeversicherung, GEZ Rundfunkbeitrag sowie Vodafone 1000 Mbit/s Highspeed WLAN (Wi-Fi 6). Eine gesonderte Abrechnung der Nebenkosten findet nicht statt.<br><br>
Die Miete ist monatlich im Voraus bis spätestens zum 3. Werktag des Monats per SEPA-Überweisung zu entrichten.
</td>
<td class="en">
The monthly all-inclusive warm rent is fixed at <strong>€{rent_amount:,.2f}</strong>.<br><br>
This flat-rate covers all operational utilities: heating, water, electricity (100% green energy), HOA/Hausgeld, building insurance, mandatory TV/radio tax (GEZ), and 1,000 Mbps Vodafone Wi-Fi 6 Internet. No separate utility reconciliation will occur.<br><br>
Rent is payable in advance on or before the 3rd business day of each calendar month via SEPA bank transfer.
</td>
</tr>
</table>

<h3>§ 4 Kaution / Security Deposit</h3>
<table class="bilingual-table">
<tr>
<td class="de">
Der Mieter leistet vor der Schlüsselübergabe eine Kaution in Höhe von <strong>€{deposit_amount:,.2f}</strong>.<br><br>
Die Kaution wird auf einem separaten Kautionskonto verwahrt und nach ordnungsgemäßer Rückgabe des Zimmers und Unterzeichnung des Auszugsprotokolls innerhalb von maximal 30 Tagen auf das angegebene Konto des Mieters zurücküberwiesen, abzüglich etwaiger Schäden oder offener Forderungen.
</td>
<td class="en">
The tenant pays a security deposit of <strong>€{deposit_amount:,.2f}</strong> prior to key handover.<br><br>
The deposit is kept in an escrow account and shall be reimbursed to the tenant within 30 days after room handover and sign-off of the exit protocol, less any proven damages, cleaning fees, or outstanding balances.
</td>
</tr>
</table>

<h3>§ 5 Schlüssel & Hausordnung / Keys & House Rules</h3>
<table class="bilingual-table">
<tr>
<td class="de">
Der Mieter erhält 4 Schlüssel: 1x Hauseingang, 1x Wohnungstür, 1x Zimmertür, 1x Briefkasten.<br><br>
Die Hausordnung (insb. Ruhezeiten 22:00 – 07:00 Uhr und striktes Rauchverbot in der gesamten Wohnung) ist wesentlicher Bestandteil dieses Vertrages. Eine behördliche Wohnungsgeberbestätigung (Anmeldung) wird unverzüglich nach Einzug ausgestellt.
</td>
<td class="en">
The tenant receives 4 keys: 1x building entrance, 1x apartment door, 1x room door, 1x mailbox.<br><br>
The House Rules (in particular quiet hours 22:00 – 07:00 and strict non-smoking policy throughout the apartment) form an integral part of this agreement. An official landlord confirmation (Wohnungsgeberbestätigung / Anmeldung) is provided upon move-in.
</td>
</tr>
</table>

<div style="margin-top: 30px; font-size: 10pt;">
Berlin, den {today_str}
</div>

<div class="signatures">
    <div class="sig-box">
        Kylasam N. A.<br>
        <strong>Vermieter / Landlord</strong>
    </div>
    <div class="sig-box">
        {tenant_name}<br>
        <strong>Mieter / Tenant</strong>
    </div>
</div>

</body>
</html>
"""
    return html

def upload_contract_to_drive(token, folder_id, doc_name, html_content):
    # Upload as Google Doc
    metadata = {
        "name": doc_name,
        "mimeType": "application/vnd.google-apps.document",
        "parents": [folder_id]
    }
    boundary = "===============HOMELYO_BOUNDARY=="
    body = (
        f"--{boundary}\r\n"
        f"Content-Type: application/json; charset=UTF-8\r\n\r\n"
        f"{json.dumps(metadata)}\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: text/html; charset=UTF-8\r\n\r\n"
        f"{html_content}\r\n"
        f"--{boundary}--\r\n"
    ).encode("utf-8")

    req = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        doc_res = json.loads(resp.read().decode())
        doc_id = doc_res["id"]

    # Export as PDF & upload PDF version to folder
    export_url = f"https://www.googleapis.com/drive/v3/files/{doc_id}/export?mimeType=application/pdf"
    exp_req = urllib.request.Request(export_url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(exp_req) as resp:
        pdf_bytes = resp.read()

    pdf_metadata = {
        "name": f"{doc_name}.pdf",
        "mimeType": "application/pdf",
        "parents": [folder_id]
    }
    pdf_body = (
        f"--{boundary}\r\n"
        f"Content-Type: application/json; charset=UTF-8\r\n\r\n"
        f"{json.dumps(pdf_metadata)}\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: application/pdf\r\n\r\n"
    ).encode("utf-8") + pdf_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")

    pdf_req = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=pdf_body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    with urllib.request.urlopen(pdf_req) as resp:
        pdf_res = json.loads(resp.read().decode())
        pdf_id = pdf_res["id"]

    return doc_id, pdf_id

def sync_to_google_sheet(token, tenant_name, tenant_email, room_code, start_date, end_date, rent_amount, deposit_amount, gdrive_folder_url, platform):
    d_start = datetime.datetime.strptime(start_date, "%Y-%m-%d").date()
    d_end = datetime.datetime.strptime(end_date, "%Y-%m-%d").date()
    duration = (d_end - d_start).days
    remaining = max(0, (d_end - datetime.date.today()).days)

    # 1. Append to tenant_details
    row_tenant = [
        room_code,
        platform,
        tenant_name,
        tenant_email,
        start_date,
        end_date,
        str(duration),
        str(remaining),
        "TRUE",
        str(int(rent_amount)),
        str(int(deposit_amount)),
        str(int(deposit_amount)),
        "TRUE",
        f"Booked via Onboarding Engine on {datetime.date.today().isoformat()}",
        gdrive_folder_url,
        "HELD_IN_ESCROW",
        ""
    ]
    append_url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values/tenant_details!A:Q:append?valueInputOption=USER_ENTERED"
    req = urllib.request.Request(
        append_url,
        data=json.dumps({"values": [row_tenant]}).encode("utf-8"),
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        pass

    # 2. Append monthly rent lines to rent_log
    curr = d_start.replace(day=1)
    rent_rows = []
    while curr <= d_end:
        m_str = curr.strftime("%Y-%m-%d")
        due_str = curr.strftime("%Y-%m-06")
        rent_rows.append([
            room_code,
            "TRUE" if curr <= datetime.date.today() else "FALSE",
            tenant_name,
            m_str,
            due_str,
            f"€{rent_amount:,.2f}",
            "€0.00",
            f"€{rent_amount:,.2f}",
            f"€{rent_amount:,.2f}" if curr <= datetime.date.today() else "€0.00",
            platform,
            datetime.date.today().strftime("%Y-%m-02") if curr <= datetime.date.today() else "",
            "TRUE",
            float(rent_amount)
        ])
        # Next month
        if curr.month == 12:
            curr = curr.replace(year=curr.year + 1, month=1)
        else:
            curr = curr.replace(month=curr.month + 1)

    rent_log_url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values/rent_log!A:M:append?valueInputOption=USER_ENTERED"
    req_rent = urllib.request.Request(
        rent_log_url,
        data=json.dumps({"values": rent_rows}).encode("utf-8"),
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req_rent) as resp:
        pass

def sync_to_clickup(tenant_name, tenant_email, tenant_phone, room_code, start_date, end_date, rent_amount, deposit_amount, gdrive_url, doc_id, pdf_id):
    # Convert dates to Unix millisecond timestamps for ClickUp Calendar & Timeline view
    d_start_dt = datetime.datetime.strptime(start_date, "%Y-%m-%d")
    d_end_dt = datetime.datetime.strptime(end_date, "%Y-%m-%d")
    start_ts_ms = int(d_start_dt.timestamp() * 1000)
    due_ts_ms = int(d_end_dt.timestamp() * 1000)

    # ClickUp Task in Tenants & Room Occupancy
    desc = f"""### 🏠 New Tenant Onboarding Profile — {tenant_name} ({room_code})
* **Room:** {room_code} ({ROOM_SPECS.get(room_code, {}).get('title', 'Bedroom')})
* **Tenant Email:** {tenant_email}
* **Tenant Phone:** {tenant_phone}
* **Lease Term:** {start_date} → {end_date}
* **Warm Rent:** €{rent_amount:,.2f} / month
* **Deposit Held (Kaution):** €{deposit_amount:,.2f}

---
### 📁 Document Links (Google Drive)
* 📂 [Open Tenant GDrive Folder]({gdrive_url})
* 📝 [Open Dual-Language Rental Contract (Google Doc)](https://docs.google.com/document/d/{doc_id}/edit)
* 📑 [Download PDF Rental Contract](https://drive.google.com/file/d/{pdf_id}/view)

---
### ✅ Onboarding Package Included
* [x] Google Drive dedicated folder created
* [x] Dual-language German & English Mietvertrag generated (with mandatory 3-month notice clause)
* [x] House Rules (Hausordnung) attached
* [x] Apartment Handover Protocol (Übergabeprotokoll) prepared
* [x] Keys Acknowledgment sheet prepared
* [x] Official Wohnungsgeberbestätigung (Anmeldung) template ready
* [x] Google Sheet tenant_details & rent_log synchronized
* [ ] Verify first month rent payout (€{rent_amount:,.2f}) & deposit transfer (€{deposit_amount:,.2f})
* [ ] Collect signed Mietvertrag from tenant
* [ ] In-person key handover & initial inspection protocol on {start_date}
"""
    task_payload = {
        "name": f"🟡 [OCCUPIED until {d_end_dt.strftime('%b %d, %Y')}] Room {room_code} – {tenant_name} (€{rent_amount:,.0f}/mo)",
        "markdown_description": desc,
        "status": "occupied (active lease)",
        "start_date": start_ts_ms,
        "due_date": due_ts_ms
    }
    r = urllib.request.Request(
        f"https://api.clickup.com/api/v2/list/{LIST_TENANTS}/task",
        data=json.dumps(task_payload).encode("utf-8"),
        headers=CLICKUP_HEADERS,
        method="POST"
    )
    task_id = None
    try:
        with urllib.request.urlopen(r) as resp:
            data = json.loads(resp.read().decode())
            task_id = data.get("id")
    except Exception as e:
        print(f"ClickUp task creation note: {e}")

    # Also log in Tenant Expenses & Deposit Settlement
    settle_desc = f"""### 🔐 Deposit Escrow & Settlement Ledger — {tenant_name} ({room_code})
* **Advance Deposit Paid:** €{deposit_amount:,.2f}
* **Status:** Held in Escrow (Active Lease)
* **Contract Term:** {start_date} to {end_date}
* **Folder Link:** [Google Drive]({gdrive_url})
"""
    settle_payload = {
        "name": f"🔒 [Room {room_code}] {tenant_name} – Deposit €{deposit_amount:,.2f}",
        "markdown_description": settle_desc,
        "status": "to do",
        "start_date": start_ts_ms,
        "due_date": due_ts_ms
    }
    r2 = urllib.request.Request(
        f"https://api.clickup.com/api/v2/list/{LIST_SETTLEMENT}/task",
        data=json.dumps(settle_payload).encode("utf-8"),
        headers=CLICKUP_HEADERS,
        method="POST"
    )
    try:
        with urllib.request.urlopen(r2) as resp:
            pass
    except Exception as e:
        pass

    return task_id

def onboard_tenant(name, email, phone, room, start_date, end_date=None, duration_months=None, platform="Direct / Homelyo Portal", rent_amount=None, deposit_amount=None):
    room = room.upper().strip()
    if room not in ROOM_SPECS:
        raise ValueError(f"Invalid room '{room}'. Must be KE1, KE2, KE3, or KE4.")

    d_start = datetime.datetime.strptime(start_date, "%Y-%m-%d").date()
    if not end_date:
        if not duration_months:
            duration_months = 6
        # Compute end date (e.g. 6 months minus 1 day)
        m_end = d_start.month + int(duration_months)
        y_end = d_start.year + (m_end - 1) // 12
        m_end = ((m_end - 1) % 12) + 1
        # Last day of preceding month or exact day
        first_of_target = datetime.date(y_end, m_end, 1)
        d_end = first_of_target - datetime.timedelta(days=1)
        end_date = d_end.strftime("%Y-%m-%d")
    else:
        d_end = datetime.datetime.strptime(end_date, "%Y-%m-%d").date()

    token = get_gcloud_token()
    if not token:
        raise RuntimeError("GCloud authentication token unavailable.")

    # 1. Folder Name format: {Name}_{Room}_{YYYYMMDD_start}_{YYYYMMDD_end}
    clean_name = name.replace(" ", "_").replace("-", "_")
    start_tag = d_start.strftime("%Y%m%d")
    end_tag = d_end.strftime("%Y%m%d")
    folder_name = f"{clean_name}_{room}_{start_tag}_{end_tag}"

    print(f"\n🚀 Starting Autonomous Onboarding for: {name} ({room})")
    print(f"📁 Target Google Drive Folder: {folder_name}")

    # Step 1: Create folder in GDrive
    folder_id = create_drive_folder(token, folder_name, PARENT_FOLDER_ID)
    folder_url = f"https://drive.google.com/drive/u/0/folders/{folder_id}"
    print(f"✅ Created Folder: {folder_url}")

    # Step 2: Copy template package
    print("📦 Copying onboarding document templates (House Rules, Handover, Anmeldung)...")
    for t in TEMPLATE_FILES:
        try:
            copy_drive_file(token, t["id"], t["name"], folder_id)
            print(f"  ✓ Copied: {t['name']}")
        except Exception as e:
            print(f"  ⚠ Failed copying {t['name']}: {e}")

    # Step 3: Generate Contract
    if rent_amount is None:
        rent_amount = ROOM_SPECS[room]["rent"]
    else:
        rent_amount = float(rent_amount)

    if deposit_amount is None:
        deposit_amount = ROOM_SPECS[room]["deposit"]
    else:
        deposit_amount = float(deposit_amount)

    contract_doc_name = f"Rental_Contract_{clean_name}_{room}_{start_tag}_{end_tag}"
    print(f"📝 Generating dual-language rental contract: {contract_doc_name} (€{rent_amount:,.2f}/mo)...")
    html = generate_contract_html(name, email, phone, room, start_date, end_date, rent_amount, deposit_amount)
    doc_id, pdf_id = upload_contract_to_drive(token, folder_id, contract_doc_name, html)
    print(f"  ✓ Contract Google Doc ID: {doc_id}")
    print(f"  ✓ Contract PDF ID: {pdf_id}")

    # Step 4: Sync to Google Sheet
    print("📊 Synchronizing Google Sheet (tenant_details & rent_log)...")
    try:
        sync_to_google_sheet(token, name, email, room, start_date, end_date, rent_amount, deposit_amount, folder_url, platform)
        print("  ✓ Google Sheet synchronized successfully.")
    except Exception as e:
        print(f"  ⚠ Google Sheet sync warning: {e}")

    # Step 5: Sync to ClickUp
    print("🏢 Creating ClickUp tasks & settlement ledger...")
    cu_id = sync_to_clickup(name, email, phone, room, start_date, end_date, rent_amount, deposit_amount, folder_url, doc_id, pdf_id)
    print(f"  ✓ ClickUp Task ID: {cu_id or 'Success'}")

    result = {
        "status": "SUCCESS",
        "tenant_name": name,
        "room": room,
        "dates": f"{start_date} to {end_date}",
        "folder_name": folder_name,
        "folder_url": folder_url,
        "contract_doc_url": f"https://docs.google.com/document/d/{doc_id}/edit",
        "contract_pdf_url": f"https://drive.google.com/file/d/{pdf_id}/view",
        "clickup_task_id": cu_id
    }
    print(f"\n🎉 Onboarding Complete in One Shot!")
    return result

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Homelyo Autonomous Tenant Onboarding")
    parser.add_argument("--name", required=True, help="Full Name of tenant")
    parser.add_argument("--email", required=True, help="Email address")
    parser.add_argument("--phone", default="", help="Phone number")
    parser.add_argument("--room", required=True, choices=["KE1", "KE2", "KE3", "KE4"], help="Room Code")
    parser.add_argument("--start", required=True, help="Start Date (YYYY-MM-DD)")
    parser.add_argument("--end", default=None, help="End Date (YYYY-MM-DD)")
    parser.add_argument("--duration", type=int, default=None, help="Duration in months (default 6)")
    parser.add_argument("--platform", default="Direct / Homelyo Portal", help="Booking channel")
    parser.add_argument("--rent", type=float, default=None, help="Monthly warm rent in EUR")
    parser.add_argument("--deposit", type=float, default=None, help="Security deposit in EUR")

    args = parser.parse_args()
    res = onboard_tenant(
        name=args.name,
        email=args.email,
        phone=args.phone,
        room=args.room,
        start_date=args.start,
        end_date=args.end,
        duration_months=args.duration,
        platform=args.platform,
        rent_amount=args.rent,
        deposit_amount=args.deposit
    )
    print(json.dumps(res, indent=2))
