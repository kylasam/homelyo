#!/usr/bin/env python3
"""
Homelyo Online Booking Form Simulation & Automated Contract Generation
======================================================================
Simulates an incoming tenant booking application submitted via the online
form (Fillout / Website webhook) and executes the end-to-end automated
contract generation pipeline:

1. Ingests mock online form submission payload
2. Extracts tenant identity, room selection, and stay parameters
3. Applies Homelyo co-living business rules & room specifications
4. Generates official 7-page bilingual agreement:
   - Microsoft Word (.docx)
   - Interactive HTML (.html)
   - High-Fidelity PDF (.pdf) via Google Chrome Headless
5. Verifies all merged rules (Internet, 100€ minor repairs, advance payment,
   4th of month deadline + 5€ fee, room details, quiet hours & kitchen etiquette).
"""

import sys
import os
import json
import datetime
import subprocess

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from build_contract_templates import (
    render_contract_data,
    render_contract_docx,
    render_contract_html,
    convert_html_to_pdf,
    ROOM_SPECS
)

def simulate_form_submission(form_payload):
    print("=" * 70)
    print("🌐 STEP 1: INCOMING ONLINE BOOKING FORM SUBMISSION RECEIVED")
    print("=" * 70)
    print(json.dumps(form_payload, indent=2, ensure_ascii=False))

    # Extract fields from form
    name = form_payload.get("full_name") or form_payload.get("name")
    email = form_payload.get("email")
    phone = form_payload.get("phone")
    room_code = form_payload.get("room_selected", "KE2")
    start_date = form_payload.get("move_in_date")
    end_date = form_payload.get("move_out_date")
    address_de = form_payload.get("address_de") or form_payload.get("current_address")
    address_en = form_payload.get("address_en") or form_payload.get("current_address")
    reason_de = form_payload.get("reason_stay_de")
    reason_en = form_payload.get("reason_stay_en")

    print("\n" + "=" * 70)
    print("⚙️  STEP 2: PARAMETER EXTRACTION & ROOM SPECIFICATION MATCHING")
    print("=" * 70)
    spec = ROOM_SPECS.get(room_code, ROOM_SPECS["KE2"])
    print(f"  • Tenant Name:      {name}")
    print(f"  • Email:            {email}")
    print(f"  • Phone:            {phone}")
    print(f"  • Room Code:        {room_code} -> {spec.get('title_en')} ({spec.get('title_de')})")
    print(f"  • Room Size:        {spec.get('size')} m² on Floor {spec.get('floor')}")
    print(f"  • Rental Period:    {start_date} to {end_date}")
    print(f"  • Monthly Flatrate: €{spec.get('rent')} (utilities + WLAN included)")
    print(f"  • Security Deposit: €{spec.get('deposit')}")

    print("\n" + "=" * 70)
    print("📄 STEP 3: EXECUTING 7-PAGE CONTRACT GENERATION PIPELINE")
    print("=" * 70)
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "contracts_generated")
    os.makedirs(output_dir, exist_ok=True)

    clean_name = "".join(c for c in name if c.isalnum() or c in (" ", "_", "-")).strip().replace(" ", "_")
    base_name = f"Rental_Contract_{clean_name}_{room_code}_{start_date.replace('-', '')}_{end_date.replace('-', '')}"

    docx_path = os.path.join(output_dir, f"{base_name}.docx")
    html_path = os.path.join(output_dir, f"{base_name}.html")
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")

    # 1. Render data dictionary
    data = render_contract_data(
        tenant_name=name,
        tenant_email=email,
        tenant_phone=phone,
        room_code=room_code,
        start_date=start_date,
        end_date=end_date,
        rent_amount=spec.get("rent"),
        deposit_amount=spec.get("deposit"),
        tenant_address_de=address_de,
        tenant_address_en=address_en,
        reason_de=reason_de,
        reason_en=reason_en
    )

    # 2. Generate Word Document (.docx)
    render_contract_docx(data, docx_path)
    print(f"  ✓ [1/3] Word Document created: {docx_path}")

    # 3. Generate HTML (.html)
    html_content = render_contract_html(data)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"  ✓ [2/3] Master HTML created:  {html_path}")

    # 4. Compile High-Fidelity PDF via Headless Chrome (.pdf)
    convert_html_to_pdf(html_path, pdf_path)
    print(f"  ✓ [3/3] High-Res PDF created: {pdf_path}")

    print("\n" + "=" * 70)
    print("🔍 STEP 4: VERIFYING MERGED RULES IN GENERATED CONTRACT")
    print("=" * 70)
    import pymupdf
    doc = pymupdf.open(pdf_path)
    print(f"  • Total Document Page Count: {len(doc)} pages (Target: 7)")
    
    full_pdf_text = ""
    for i, p in enumerate(doc):
        full_pdf_text += f"\n--- PAGE {i+1} ---\n" + p.get_text("text")

    rule_checks = [
        ("Homelyo Logo Header", "Mietvertrag" in full_pdf_text and "Rental Agreement" in full_pdf_text),
        ("Room Details Included", room_code in full_pdf_text and spec.get("size") in full_pdf_text),
        ("Advance Payment Before Move-In", "vor dem Einzugstermin" in full_pdf_text or "before the move-in date" in full_pdf_text),
        ("Rent Due Before 4th + 5€ Fee", "vor dem 4." in full_pdf_text and "5,00" in full_pdf_text),
        ("Internet / WLAN / Torrent Policy", "File-Sharing" in full_pdf_text or "peer-to-peer" in full_pdf_text),
        ("Minor Repairs < 100 EUR Clause", "100,00" in full_pdf_text or "100.00" in full_pdf_text),
        ("Light Premature Termination", "Nachmieter" in full_pdf_text or "replacement tenant" in full_pdf_text),
        ("Common Area Etiquette & Kitchen", "nach jeder Mahlzeit" in full_pdf_text or "immediately after each use" in full_pdf_text),
        ("Quiet Hours & 5% Increase", "20:00" in full_pdf_text and "5 %" in full_pdf_text)
    ]

    for label, status in rule_checks:
        symbol = "✓ PASS" if status else "✗ FAIL"
        print(f"  {symbol}: {label}")

    print("\n" + "=" * 70)
    print("🎉 STEP 5: AUTOMATED ONBOARDING PIPELINE EXECUTION COMPLETE")
    print("=" * 70)
    print(f"Ready for instant tenant signing & dispatch:")
    print(f"• PDF File:  {pdf_path}")
    print(f"• HTML File: {html_path}")
    print(f"• DOCX File: {docx_path}\n")

    return {
        "status": "success",
        "pdf_path": pdf_path,
        "html_path": html_path,
        "docx_path": docx_path,
        "page_count": len(doc)
    }

def main():
    # Simulated form submission payload as would arrive from Fillout / web booking
    mock_payload = {
        "form_id": "homelyo_room_reservation_v2",
        "submission_id": "sub_20261026_9821a",
        "submitted_at": "2026-10-26T21:40:00Z",
        "full_name": "Priya Sharma",
        "email": "priya.sharma@example.com",
        "phone": "+49 176 98765432",
        "room_selected": "KE2",
        "move_in_date": "2026-05-01",
        "move_out_date": "2026-10-31",
        "current_address": "Indira Nagar 100 Feet Rd, 560038 Bengaluru, India",
        "address_de": "Indira Nagar 100 Feet Rd, 560038 Bengaluru, Indien",
        "address_en": "Indira Nagar 100 Feet Rd, 560038 Bengaluru, India",
        "reason_stay_de": "Aufgrund eines Forschungsaufenthalts / Masterstudiums an der Freien Universität Berlin",
        "reason_stay_en": "Due to a research stay / master's degree program at Freie Universität Berlin"
    }

    simulate_form_submission(mock_payload)

if __name__ == "__main__":
    main()
