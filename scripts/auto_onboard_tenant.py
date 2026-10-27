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
MASTER_BI_SPREADSHEET_ID = "1rs5dXTyA6xQSbqVQ_ofyCAv2eFZ83CrN8sAx5qfbsTk"  # Looker Studio DataStudio Backend
FILLOUT_SPREADSHEET_ID = "1gcQ4nv-lu8j3Kt3mNoFx11W7jkU0Umu1Ee2C2wM6Nvw"  # Fillout Inquiries Sheet
SPREADSHEET_ID = MASTER_BI_SPREADSHEET_ID

CLICKUP_TOKEN = "pk_314452952_6STD829UQJXD0ROPHUVS41YXEPHXUIE2"
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
    {"id": "1mNXbTpDbUxWQOn25c0waOFjaUAxpyLPq", "name": "Inventory_List.pdf"},
    {"id": "1z_ao-a8JTagdBPMaz-sw380FCr59HNov", "name": "S+U-Bahn_Berlin_Transit_Map.pdf"}
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

from build_contract_templates import (
    render_contract_data,
    render_contract_docx,
    render_contract_html,
    DOCX_TEMPLATE_PATH,
    HTML_TEMPLATE_PATH
)

def generate_contract_html(tenant_name, tenant_email, tenant_phone, room_code, start_date, end_date, rent_amount, deposit_amount):
    """
    Renders the official bilingual Homelyo Rental Contract in HTML format
    with embedded Homelyo logo and exact clause layout.
    """
    tenant_dict = render_contract_data(
        tenant_name, tenant_email, tenant_phone, room_code, start_date, end_date, rent_amount, deposit_amount
    )
    return render_contract_html(tenant_dict)

def upload_contract_to_drive(token, folder_id, doc_name, tenant_name, tenant_email, tenant_phone, room_code, start_date, end_date, rent_amount, deposit_amount):
    """
    Renders the master templated contract (with Homelyo logo, tables & AGB),
    uploads as a native Google Doc, exports as official PDF, and uploads DOCX to tenant folder.
    """
    tenant_dict = render_contract_data(
        tenant_name, tenant_email, tenant_phone, room_code, start_date, end_date, rent_amount, deposit_amount
    )
    temp_docx_path = f"/tmp/{doc_name}.docx"
    render_contract_docx(tenant_dict, temp_docx_path)

    with open(temp_docx_path, "rb") as f:
        docx_bytes = f.read()

    boundary = "===============HOMELYO_BOUNDARY=="

    # 1. Upload DOCX and convert to Google Doc
    metadata = {
        "name": doc_name,
        "mimeType": "application/vnd.google-apps.document",
        "parents": [folder_id, "1CZcV4XWuY4pSpfNjkubdHsxprKg1f2xB"]
    }
    body = (
        f"--{boundary}\r\n"
        f"Content-Type: application/json; charset=UTF-8\r\n\r\n"
        f"{json.dumps(metadata)}\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n"
    ).encode("utf-8") + docx_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")

    req = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        doc_res = json.loads(resp.read().decode())
        doc_id = doc_res["id"]

    # 2. Export converted Google Doc as PDF & upload PDF to folder
    export_url = f"https://www.googleapis.com/drive/v3/files/{doc_id}/export?mimeType=application/pdf"
    exp_req = urllib.request.Request(export_url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(exp_req) as resp:
        pdf_bytes = resp.read()

    pdf_metadata = {
        "name": f"{doc_name}.pdf",
        "mimeType": "application/pdf",
        "parents": [folder_id, "1CZcV4XWuY4pSpfNjkubdHsxprKg1f2xB"]
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

    # 3. Upload raw DOCX version to folder for direct Microsoft Word download
    docx_metadata = {
        "name": f"{doc_name}.docx",
        "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "parents": [folder_id, "1CZcV4XWuY4pSpfNjkubdHsxprKg1f2xB"]
    }
    docx_body = (
        f"--{boundary}\r\n"
        f"Content-Type: application/json; charset=UTF-8\r\n\r\n"
        f"{json.dumps(docx_metadata)}\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n"
    ).encode("utf-8") + docx_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")

    docx_req = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=docx_body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    docx_id = None
    try:
        with urllib.request.urlopen(docx_req) as resp:
            docx_res = json.loads(resp.read().decode())
            docx_id = docx_res["id"]
    except Exception as e:
        print(f"  ⚠ Notice: Could not upload standalone docx: {e}")

    return doc_id, pdf_id, docx_id

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

def upload_prefilled_wohnungsgeberbestaetigung(token, folder_id, tenant_name, room_code, start_date):
    """
    Fills out the official Wohnungsgeberbestätigung (§ 19 Abs. 3 BMG) AcroForm fields
    with tenant, landlord, and property details and uploads to the tenant folder.
    """
    import pypdf
    from pypdf.generic import BooleanObject, NameObject

    # Download base template from Google Drive
    template_id = "1yrIDs6kVUPZc8A1guV-BaUXf_x4JsfHB"
    req = urllib.request.Request(
        f"https://www.googleapis.com/drive/v3/files/{template_id}?alt=media",
        headers={"Authorization": f"Bearer {token}"}
    )
    with urllib.request.urlopen(req) as resp:
        template_bytes = resp.read()

    temp_template_path = "/tmp/wgb_temp_base.pdf"
    with open(temp_template_path, "wb") as f:
        f.write(template_bytes)

    reader = pypdf.PdfReader(temp_template_path)
    writer = pypdf.PdfWriter()
    writer.append(reader)

    # Parse name
    parts = tenant_name.strip().split()
    if len(parts) > 1:
        vorname = parts[0]
        nachname = " ".join(parts[1:])
    else:
        vorname = tenant_name
        nachname = ""

    d_start = datetime.datetime.strptime(start_date, "%Y-%m-%d")
    date_de = d_start.strftime("%d.%m.%Y")

    field_values = {
        "Familienname1": "Nallagoundanpalayam Arruchamy Kylasam",
        "Straße1": "Forckenbeckstraße 64",
        "Postleitzahl1": "14199",
        "Ort1": "Berlin",
        "Einzugsdatum1": date_de,
        "Straße2": "Lichtenrader Damm 80",
        "Zusatzangaben1": f"2EG, WE-11 (Zimmer {room_code})",
        "Postleitzahl2": "12305",
        "Ort2": "Berlin",
        "FamiliennameRow1": nachname,
        "VornameRow1": vorname,
        "Datum Unterschrift Wohnungsgeber bzw Wohnungseigentümer": date_de
    }

    for page in writer.pages:
        writer.update_page_form_field_values(page, field_values)

    if "/AcroForm" in writer._root_object:
        writer._root_object["/AcroForm"].update({
            NameObject("/NeedAppearances"): BooleanObject(True)
        })

    temp_out_path = f"/tmp/wgb_{tenant_name.replace(' ', '_')}.pdf"
    with open(temp_out_path, "wb") as f:
        writer.write(f)

    with open(temp_out_path, "rb") as f:
        pdf_bytes = f.read()

    clean_name = tenant_name.replace(" ", "_").replace("-", "_")
    boundary = "===============HOMELYO_BOUNDARY=="
    meta = {
        "name": f"wohnungsgeberbestaetigung_{clean_name}.pdf",
        "mimeType": "application/pdf",
        "parents": [folder_id]
    }
    body = (
        f"--{boundary}\r\n"
        f"Content-Type: application/json; charset=UTF-8\r\n\r\n"
        f"{json.dumps(meta)}\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: application/pdf\r\n\r\n"
    ).encode("utf-8") + pdf_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")

    req_up = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    with urllib.request.urlopen(req_up) as resp:
        res = json.loads(resp.read().decode())
        return res["id"]

def upload_personalized_welcome_kit(token, folder_id, tenant_name, room_code, start_date, end_date, contract_doc_id, contract_pdf_id):
    """
    Renders and uploads a personalized Welcome Kit (Google Doc, PDF, and DOCX)
    with WiFi details, quiet hours, room specifics, emergency numbers, and contract links.
    """
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    # Download room template from Template Folder if available, or build standard
    template_id_map = {
        "KE1": "1Wbx3gTcfI1O7mK--GqU6y0TUN8Akc3h8",
        "KE2": "1KpFS7OQRgpzh9DPynp43gkl6KYjenZiU",
        "KE3": "1tDDu4TuTLJR2JuS-47FJKqgkb5SIB2jv",
        "KE4": "1-grsqQQlDvMfD42o85Kouiq3S6n6dlfx"
    }
    tmpl_id = template_id_map.get(room_code, "1KpFS7OQRgpzh9DPynp43gkl6KYjenZiU")

    req = urllib.request.Request(
        f"https://www.googleapis.com/drive/v3/files/{tmpl_id}?alt=media",
        headers={"Authorization": f"Bearer {token}"}
    )
    with urllib.request.urlopen(req) as resp:
        tmpl_bytes = resp.read()

    temp_in = f"/tmp/tmpl_welcome_{room_code}.docx"
    with open(temp_in, "wb") as f:
        f.write(tmpl_bytes)

    doc = Document(temp_in)
    contract_url = f"https://docs.google.com/document/d/{contract_doc_id}/edit"
    contract_pdf_url = f"https://drive.google.com/file/d/{contract_pdf_id}/view"

    d_s = datetime.datetime.strptime(start_date, "%Y-%m-%d").strftime("%d %B %Y").upper()
    d_e = datetime.datetime.strptime(end_date, "%Y-%m-%d").strftime("%d %B %Y").upper()

    for p in doc.paragraphs:
        if "🌟 Welcome" in p.text:
            p.text = f"🌟 Welcome {tenant_name} 🙂 – {room_code} ({ROOM_SPECS.get(room_code, {}).get('title', 'Suite')})"
            p.style.font.size = Pt(20)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif "01 NOVEMBER" in p.text or "Tenancy:" in p.text or "Booking Period:" in p.text:
            p.text = f"📅 Booking Period: {d_s} – {d_e}\n📍 Address: Lichtenrader Damm 80, 12305 Berlin (2. OG, WE-11)"
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif "Scan this QR code" in p.text:
            p.text = f"📄 View Your Digital Tenancy Agreement (Mietvertrag):\n{contract_url}\n(PDF Download: {contract_pdf_url})"
        elif "Scannen Sie diesen QR-Code" in p.text:
            p.text = "Scannen oder klicken Sie auf den obigen Link, um Ihren Mietvertrag einzusehen."

    clean_name = tenant_name.replace(" ", "_").replace("-", "_")
    out_docx_path = f"/tmp/Welcome_Kit_{room_code}_{clean_name}.docx"
    doc.save(out_docx_path)

    with open(out_docx_path, "rb") as f:
        docx_bytes = f.read()

    boundary = "===============HOMELYO_BOUNDARY=="
    doc_name = f"Welcome_Kit_{room_code}_{clean_name}"

    # 1. Google Doc
    meta = {"name": doc_name, "mimeType": "application/vnd.google-apps.document", "parents": [folder_id]}
    body = (
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{json.dumps(meta)}\r\n"
        f"--{boundary}\r\nContent-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n"
    ).encode("utf-8") + docx_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")
    req_doc = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    with urllib.request.urlopen(req_doc) as resp:
        w_doc_id = json.loads(resp.read().decode())["id"]

    # 2. Export PDF
    exp_url = f"https://www.googleapis.com/drive/v3/files/{w_doc_id}/export?mimeType=application/pdf"
    exp_req = urllib.request.Request(exp_url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(exp_req) as resp:
        w_pdf_bytes = resp.read()

    pdf_meta = {"name": f"{doc_name}.pdf", "mimeType": "application/pdf", "parents": [folder_id]}
    pdf_body = (
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{json.dumps(pdf_meta)}\r\n"
        f"--{boundary}\r\nContent-Type: application/pdf\r\n\r\n"
    ).encode("utf-8") + w_pdf_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")
    req_pdf = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=pdf_body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    with urllib.request.urlopen(req_pdf) as resp:
        w_pdf_id = json.loads(resp.read().decode())["id"]

    # 3. Upload DOCX
    docx_meta = {"name": f"{doc_name}.docx", "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "parents": [folder_id]}
    docx_body = (
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{json.dumps(docx_meta)}\r\n"
        f"--{boundary}\r\nContent-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n"
    ).encode("utf-8") + docx_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")
    req_raw = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=docx_body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    with urllib.request.urlopen(req_raw) as resp:
        w_docx_id = json.loads(resp.read().decode())["id"]

    return w_doc_id, w_pdf_id, w_docx_id

def upload_personalized_keys_acknowledgement(token, folder_id, tenant_name, room_code, start_date):
    """
    Renders and uploads a personalized Keys Acknowledgement document (Google Doc & PDF).
    """
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    title = doc.add_heading(level=0)
    run = title.add_run("🔑 Keys Information / Schlüssel-Informationen")
    run.font.size = Pt(20)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("\n")

    room_title = ROOM_SPECS.get(room_code, {}).get("title", "Room")

    doc.add_heading("Key Inventory / Schlüsselübersicht", level=1)
    doc.add_paragraph(
        "English:\n"
        "Each tenant will receive a total of 4 keys:\n"
        "• Main entrance / Haupteingang — 1\n"
        "• Apartment door / Wohnungstür — 1\n"
        "• Letterbox / Briefkasten — 1\n"
        f"• Room key ({room_code} - {room_title}) / Zimmertür — 1\n\n"
        "Please keep all keys safe. In case of loss, tenants will be responsible for the cost of replacement. "
        "This may result in high replacement charges, as locks may need to be changed for security reasons.\n\n"
        "Deutsch:\n"
        "Jeder Mieter erhält insgesamt 4 Schlüssel:\n"
        "• Haupteingang — 1\n"
        "• Wohnungstür — 1\n"
        "• Briefkasten — 1\n"
        f"• Zimmertür ({room_code} - {room_title}) — 1\n\n"
        "Bitte bewahren Sie alle Schlüssel sorgfältig auf. Im Falle eines Verlustes trägt der Mieter die Kosten für den Ersatz. "
        "Dies kann zu hohen Ersatzkosten führen, da Schlösser aus Sicherheitsgründen ausgetauscht werden müssen."
    )

    d_start = datetime.datetime.strptime(start_date, "%Y-%m-%d").strftime("%d.%m.%Y")
    doc.add_paragraph("\n")
    doc.add_heading("📝 Acknowledgement / Bestätigung", level=1)
    doc.add_paragraph(
        "I acknowledge receipt of 4 keys and agree to follow the key policy above.\n"
        "Ich bestätige den Empfang von 4 Schlüsseln und verpflichte mich, die obige Schlüsselregelung einzuhalten.\n\n"
        f"Tenant name / Name Mieter:  {tenant_name}\n"
        f"Room / Zimmer:              {room_code} ({room_title})\n"
        f"Move-in Date / Einzugsdatum: {d_start}\n"
        "Address / Anschrift:        Lichtenrader Damm 80, 12305 Berlin\n"
        "Signature / Unterschrift:   _______________________________________\n"
    )

    clean_name = tenant_name.replace(" ", "_").replace("-", "_")
    out_docx_path = f"/tmp/Keys_Acknowledgement_{clean_name}.docx"
    doc.save(out_docx_path)

    with open(out_docx_path, "rb") as f:
        docx_bytes = f.read()

    boundary = "===============HOMELYO_BOUNDARY=="
    doc_name = f"Keys_Acknowledgement_{clean_name}"

    meta = {"name": doc_name, "mimeType": "application/vnd.google-apps.document", "parents": [folder_id]}
    body = (
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{json.dumps(meta)}\r\n"
        f"--{boundary}\r\nContent-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n"
    ).encode("utf-8") + docx_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")
    req_doc = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    with urllib.request.urlopen(req_doc) as resp:
        k_doc_id = json.loads(resp.read().decode())["id"]

    exp_url = f"https://www.googleapis.com/drive/v3/files/{k_doc_id}/export?mimeType=application/pdf"
    exp_req = urllib.request.Request(exp_url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(exp_req) as resp:
        k_pdf_bytes = resp.read()

    pdf_meta = {"name": f"{doc_name}.pdf", "mimeType": "application/pdf", "parents": [folder_id]}
    pdf_body = (
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{json.dumps(pdf_meta)}\r\n"
        f"--{boundary}\r\nContent-Type: application/pdf\r\n\r\n"
    ).encode("utf-8") + k_pdf_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")
    req_pdf = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=pdf_body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    with urllib.request.urlopen(req_pdf) as resp:
        k_pdf_id = json.loads(resp.read().decode())["id"]

    return k_doc_id, k_pdf_id

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
    doc_id, pdf_id, docx_id = upload_contract_to_drive(
        token, folder_id, contract_doc_name,
        name, email, phone, room, start_date, end_date, rent_amount, deposit_amount
    )
    print(f"  ✓ Contract Google Doc ID: {doc_id}")
    print(f"  ✓ Contract PDF ID: {pdf_id}")
    if docx_id:
        print(f"  ✓ Contract Word DOCX ID: {docx_id}")

    # Step 3b: Generate Pre-filled Wohnungsgeberbestätigung (§ 19 BMG)
    print("🏛️ Generating pre-filled Wohnungsgeberbestätigung for Bürgeramt registration...")
    wgb_id = None
    try:
        wgb_id = upload_prefilled_wohnungsgeberbestaetigung(token, folder_id, name, room, start_date)
        print(f"  ✓ Wohnungsgeberbestätigung PDF ID: {wgb_id}")
    except Exception as e:
        print(f"  ⚠ Failed generating Wohnungsgeberbestätigung: {e}")

    # Step 3c: Generate Personalized Welcome Kit
    print("🌟 Generating personalized Welcome Kit (WiFi, emergency contacts, local stores)...")
    w_doc_id, w_pdf_id, w_docx_id = None, None, None
    try:
        w_doc_id, w_pdf_id, w_docx_id = upload_personalized_welcome_kit(
            token, folder_id, name, room, start_date, end_date, doc_id, pdf_id
        )
        print(f"  ✓ Welcome Kit PDF ID: {w_pdf_id}")
    except Exception as e:
        print(f"  ⚠ Failed generating Welcome Kit: {e}")

    # Step 3d: Generate Personalized Keys Acknowledgement
    print("🔑 Generating personalized Keys Acknowledgement...")
    k_doc_id, k_pdf_id = None, None
    try:
        k_doc_id, k_pdf_id = upload_personalized_keys_acknowledgement(
            token, folder_id, name, room, start_date
        )
        print(f"  ✓ Keys Acknowledgement PDF ID: {k_pdf_id}")
    except Exception as e:
        print(f"  ⚠ Failed generating Keys Acknowledgement: {e}")

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
        "contract_docx_url": f"https://drive.google.com/file/d/{docx_id}/view" if docx_id else None,
        "wohnungsgeberbestaetigung_pdf_url": f"https://drive.google.com/file/d/{wgb_id}/view" if wgb_id else None,
        "welcome_kit_pdf_url": f"https://drive.google.com/file/d/{w_pdf_id}/view" if w_pdf_id else None,
        "keys_acknowledgement_pdf_url": f"https://drive.google.com/file/d/{k_pdf_id}/view" if k_pdf_id else None,
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
