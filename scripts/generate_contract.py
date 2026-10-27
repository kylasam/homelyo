#!/usr/bin/env python3
"""
Homelyo On-Demand Contract Generator
====================================
Generates a fully styled, bilingual (German & English) Homelyo Co-Living
Rental Contract matching the exact 7-page format of Bhargavi's signed agreement
(SIGNED_CONTRACT_67da88a52ff5e11a1f9d8ffa_20250325..pdf), complete with:
- High-resolution Homelyo logo embedded at top of Page 1 (replacing Wunderflats)
- Automatic replacement of all tenant, booking, and room details
- Dual-column bilingual layout (German on left, English on right)
- Exact 7-page structure with verified N26 and Commerzbank details
- Native local PDF generation via headless Google Chrome
- Optional direct sync/upload to Google Drive and PDF export

Usage:
    python3 generate_contract.py --name "Bhargavi Mangesh Ayare" --room KE2 --start 2025-03-25 --end 2025-09-30 --rent 639 --deposit 1200
    python3 generate_contract.py --name "Naeem Ullah Khan" --room KE2 --start 2025-11-01 --end 2026-04-30 --rent 649 --deposit 1298
"""

import sys
import os
import argparse
import datetime
import subprocess
import urllib.request
import json

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from build_contract_templates import (
    render_contract_data,
    render_contract_docx,
    render_contract_html,
    convert_html_to_pdf,
    ROOM_SPECS,
    DOCX_TEMPLATE_PATH,
    HTML_TEMPLATE_PATH
)

def get_gcloud_token():
    try:
        return subprocess.check_output(["gcloud", "auth", "print-access-token"]).decode().strip()
    except Exception as e:
        print(f"⚠ Could not fetch gcloud token: {e}")
        return None

def generate_local_contract(
    name,
    email="",
    phone="",
    room="KE2",
    start_date="2025-03-25",
    end_date="2025-09-30",
    rent=None,
    deposit=None,
    address_de=None,
    address_en=None,
    reason_de=None,
    reason_en=None,
    individual_de=None,
    individual_en=None,
    output_dir=None
):
    if not output_dir:
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "contracts_generated")
    os.makedirs(output_dir, exist_ok=True)

    clean_name = "".join(c for c in name if c.isalnum() or c in (" ", "_", "-")).strip().replace(" ", "_")
    base_filename = f"Rental_Contract_{clean_name}_{room}_{start_date.replace('-', '')}_{end_date.replace('-', '')}"

    data = render_contract_data(
        tenant_name=name,
        tenant_email=email,
        tenant_phone=phone,
        room_code=room,
        start_date=start_date,
        end_date=end_date,
        rent_amount=rent,
        deposit_amount=deposit,
        tenant_address_de=address_de,
        tenant_address_en=address_en,
        reason_de=reason_de,
        reason_en=reason_en,
        individual_de=individual_de,
        individual_en=individual_en
    )

    docx_output = os.path.join(output_dir, f"{base_filename}.docx")
    html_output = os.path.join(output_dir, f"{base_filename}.html")
    pdf_output = os.path.join(output_dir, f"{base_filename}.pdf")

    # 1. Render DOCX
    render_contract_docx(data, docx_output)
    print(f"  ✓ Word Document generated: {docx_output}")

    # 2. Render HTML
    html_content = render_contract_html(data)
    with open(html_output, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"  ✓ HTML Preview generated:  {html_output}")

    # 3. Export High-Fidelity PDF via Headless Chrome
    try:
        convert_html_to_pdf(html_output, pdf_output)
        print(f"  ✓ PDF Contract generated:  {pdf_output}")
    except Exception as e:
        print(f"  ⚠ Could not render local PDF directly: {e}")

    return {
        "base_filename": base_filename,
        "docx_path": docx_output,
        "html_path": html_output,
        "pdf_path": pdf_output if os.path.exists(pdf_output) else None,
        "data": data
    }

def upload_and_export_pdf(token, folder_id, doc_name, docx_path, pdf_path=None):
    with open(docx_path, "rb") as f:
        docx_bytes = f.read()

    boundary = "===============HOMELYO_BOUNDARY=="

    metadata = {
        "name": doc_name,
        "mimeType": "application/vnd.google-apps.document"
    }
    if folder_id:
        metadata["parents"] = [folder_id]

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

    if pdf_path and os.path.exists(pdf_path):
        with open(pdf_path, "rb") as pf:
            pdf_bytes = pf.read()
    else:
        export_url = f"https://www.googleapis.com/drive/v3/files/{doc_id}/export?mimeType=application/pdf"
        exp_req = urllib.request.Request(export_url, headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(exp_req) as resp:
            pdf_bytes = resp.read()

    pdf_metadata = {
        "name": f"{doc_name}.pdf",
        "mimeType": "application/pdf"
    }
    if folder_id:
        pdf_metadata["parents"] = [folder_id]

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

    docx_metadata = {
        "name": f"{doc_name}.docx",
        "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }
    if folder_id:
        docx_metadata["parents"] = [folder_id]

    raw_body = (
        f"--{boundary}\r\n"
        f"Content-Type: application/json; charset=UTF-8\r\n\r\n"
        f"{json.dumps(docx_metadata)}\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n"
    ).encode("utf-8") + docx_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")

    raw_req = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=raw_body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method="POST"
    )
    try:
        with urllib.request.urlopen(raw_req) as resp:
            pass
    except Exception:
        pass

    return doc_id, pdf_id, pdf_bytes

def main():
    parser = argparse.ArgumentParser(description="Homelyo Contract Generator")
    parser.add_argument("--name", required=True, help="Full Tenant Name (e.g., 'Bhargavi Mangesh Ayare' or 'Naeem Ullah Khan')")
    parser.add_argument("--email", default="", help="Tenant Email")
    parser.add_argument("--phone", default="", help="Tenant Phone Number")
    parser.add_argument("--room", default="KE2", choices=["KE1", "KE2", "KE3", "KE4"], help="Room ID (KE1 to KE4)")
    parser.add_argument("--start", default="2025-03-25", help="Start Date (YYYY-MM-DD)")
    parser.add_argument("--end", default="2025-09-30", help="End Date (YYYY-MM-DD)")
    parser.add_argument("--rent", type=float, default=None, help="Monthly Rent in EUR (default: standard room rate)")
    parser.add_argument("--deposit", type=float, default=None, help="Security Deposit in EUR (default: 2 months rent)")
    parser.add_argument("--address-de", default=None, help="Tenant Address in German")
    parser.add_argument("--address-en", default=None, help="Tenant Address in English")
    parser.add_argument("--reason-de", default=None, help="Reason for temporary stay in German")
    parser.add_argument("--reason-en", default=None, help="Reason for temporary stay in English")
    parser.add_argument("--individual-de", default=None, help="Individual stipulations in German")
    parser.add_argument("--individual-en", default=None, help="Individual stipulations in English")
    parser.add_argument("--output-dir", default=None, help="Local directory to write output files")
    parser.add_argument("--upload-to-drive", action="store_true", help="Upload to Google Drive & generate live Google Doc & PDF")
    parser.add_argument("--drive-folder-id", default=None, help="Target Google Drive folder ID")

    args = parser.parse_args()

    print(f"\n========================================================")
    print(f"🏠 HOMELYO CONTRACT GENERATOR (7-PAGE BILINGUAL FORMAT)")
    print(f"========================================================")
    print(f"👤 Tenant: {args.name}")
    print(f"🛏️  Room:   {args.room} ({ROOM_SPECS.get(args.room, {}).get('title', '')})")
    print(f"📅 Dates:  {args.start} to {args.end}")
    
    gen_result = generate_local_contract(
        name=args.name,
        email=args.email,
        phone=args.phone,
        room=args.room,
        start_date=args.start,
        end_date=args.end,
        rent=args.rent,
        deposit=args.deposit,
        address_de=args.address_de,
        address_en=args.address_en,
        reason_de=args.reason_de,
        reason_en=args.reason_en,
        individual_de=args.individual_de,
        individual_en=args.individual_en,
        output_dir=args.output_dir
    )

    if args.upload_to_drive:
        print("\n☁️ Uploading to Google Drive & rendering official PDF...")
        token = get_gcloud_token()
        if token:
            doc_id, pdf_id, pdf_bytes = upload_and_export_pdf(
                token, args.drive_folder_id, gen_result["base_filename"], gen_result["docx_path"], gen_result.get("pdf_path")
            )
            local_pdf = gen_result["docx_path"].replace(".docx", ".pdf")
            with open(local_pdf, "wb") as f:
                f.write(pdf_bytes)
            print(f"  ✓ Live Google Doc: https://docs.google.com/document/d/{doc_id}/edit")
            print(f"  ✓ Live PDF File:   https://drive.google.com/file/d/{pdf_id}/view")
            print(f"  ✓ Local PDF Saved: {local_pdf}")
        else:
            print("  ⚠ Could not upload: missing or expired gcloud authentication token.")

    print(f"\n🎉 Contract Generation Completed Successfully!\n")

if __name__ == "__main__":
    main()
