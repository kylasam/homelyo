#!/usr/bin/env python3
"""
Homelyo Executive Welcome Kit & Printable Desk QR Generator
----------------------------------------------------------
1. Completely redesigns the Welcome Kit PDF & Word Document (.docx + Google Doc)
   into a luxury executive editorial layout ( eliminating narrow vertical text boxes,
   overlapping shapes, or cramped columns).
2. Generates a dedicated Printable Desk QR Sheet (Word .docx, Google Doc & PDF)
   under Umakant's folder to print and paste on his desk in Room KE1.
"""

import os
import io
import json
import base64
import datetime
import subprocess
import urllib.request
import qrcode
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

ROOM_SPECS = {
    "KE1": {"title": "Master Bedroom Suite", "size": 21, "rent": 609.00, "deposit": 1000.00},
    "KE2": {"title": "Private Balcony Suite", "size": 24, "rent": 629.00, "deposit": 1258.00},
    "KE3": {"title": "Quiet Garden View Suite", "size": 12, "rent": 619.00, "deposit": 1238.00},
    "KE4": {"title": "Cozy Calm Suite", "size": 11, "rent": 569.00, "deposit": 1138.00}
}

LINKTREE_URLS = {
    "KE1": "https://linktr.ee/lichtenraderdamm_80_ke_1",
    "KE2": "https://linktr.ee/lichtenraderdamm_80_ke_2",
    "KE3": "https://linktr.ee/lichtenraderdamm_80_ke_3",
    "KE4": "https://linktr.ee/lichtenraderdamm_80_ke_4"
}


def get_gcloud_token():
    return subprocess.check_output(["gcloud", "auth", "print-access-token"]).decode().strip()


def make_qr_png(url, out_path, fill_color="#0F172A", back_color="#FFFFFF"):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=14,
        border=3,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color=fill_color, back_color=back_color)
    img.save(out_path)
    return out_path


def generate_qr_data_uri(url, fill_color="#0F172A"):
    tmp_p = f"/tmp/qr_exec_{abs(hash(url + fill_color))}.png"
    make_qr_png(url, tmp_p, fill_color=fill_color)
    with open(tmp_p, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    return f"data:image/png;base64,{b64}"


def add_hyperlink(paragraph, url, text, color_hex="0284C7", bold=True, font_size=9.5):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    c = OxmlElement("w:color")
    c.set(qn("w:val"), color_hex)
    rPr.append(c)

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)

    if bold:
        b = OxmlElement("w:b")
        rPr.append(b)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(font_size * 2)))
    rPr.append(sz)

    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def build_welcome_kit_docx(
    tenant_name,
    room_code,
    start_date,
    end_date,
    rent_amount,
    deposit_amount,
    drive_folder_url,
    contract_pdf_url,
    anmeldung_pdf_url,
    keys_pdf_url,
    out_docx_path
):
    room_info = ROOM_SPECS.get(room_code, {"title": "Private Suite", "size": 21})
    linktree_url = LINKTREE_URLS.get(room_code, "https://linktr.ee/lichtenraderdamm_80_ke_1")
    master_hub_url = "https://linktr.ee/lichtenrder_damm_80"
    service_url = "https://www.homelyo.eu/assistant/"

    qr_lt_path = f"/tmp/qr_lt_{room_code}.png"
    qr_vault_path = f"/tmp/qr_vault_{room_code}.png"
    make_qr_png(linktree_url, qr_lt_path, fill_color="#0F172A")
    make_qr_png(drive_folder_url, qr_vault_path, fill_color="#065F46")

    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor(30, 41, 59)

    # Top Header Table (Full Width Banner)
    hdr_tbl = doc.add_table(rows=1, cols=1)
    hdr_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cell = hdr_tbl.cell(0, 0)
    set_cell_shading(hdr_cell, "0F172A")
    set_cell_margins(hdr_cell, top=180, bottom=180, left=220, right=220)

    p_hdr = hdr_cell.paragraphs[0]
    r_brand = p_hdr.add_run("HOMELYO LIVING  ·  BERLIN LICHTENRADE CO-LIVING RESIDENCE\n")
    r_brand.font.size = Pt(8.5)
    r_brand.font.bold = True
    r_brand.font.color.rgb = RGBColor(52, 211, 153)

    r_title = p_hdr.add_run(f"Welcome Home, {tenant_name} — Room {room_code}\n")
    r_title.font.size = Pt(17)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(255, 255, 255)

    r_sub = p_hdr.add_run(
        f"Lichtenrader Damm 80, 12305 Berlin (2. OG, WE-11)  |  {room_info['title']} ({room_info['size']} m²)  |  "
        f"Tenancy: {start_date} – {end_date}  |  All-Inclusive Warm Rent: €{rent_amount:,.2f}/mo"
    )
    r_sub.font.size = Pt(9.5)
    r_sub.font.color.rgb = RGBColor(203, 213, 225)

    doc.add_paragraph()

    # Section 1: Wide Horizontal QR Banner Cards (Side-by-Side QR Left + Details Right)
    p_s1 = doc.add_paragraph()
    r_s1 = p_s1.add_run("1. YOUR DIGITAL RESIDENT PORTAL & PERSONAL CLOUD VAULT")
    r_s1.font.size = Pt(11)
    r_s1.font.bold = True
    r_s1.font.color.rgb = RGBColor(15, 23, 42)

    qr_tbl = doc.add_table(rows=2, cols=2)
    qr_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Row 0: Personal Cloud Vault (My_docs)
    c_img1 = qr_tbl.cell(0, 0)
    c_txt1 = qr_tbl.cell(0, 1)
    set_cell_shading(c_img1, "ECFDF5")
    set_cell_shading(c_txt1, "ECFDF5")
    set_cell_margins(c_img1, 140, 140, 140, 140)
    set_cell_margins(c_txt1, 140, 140, 180, 180)
    c_img1.width = Inches(1.7)
    c_txt1.width = Inches(5.4)

    p_i1 = c_img1.paragraphs[0]
    p_i1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_i1.add_run().add_picture(qr_vault_path, width=Inches(1.35))

    p_t1 = c_txt1.paragraphs[0]
    r_tag1 = p_t1.add_run("PRIVATE TENANT DOCUMENT VAULT (PERSONAL ACCESS ONLY)\n")
    r_tag1.font.size = Pt(8.5)
    r_tag1.font.bold = True
    r_tag1.font.color.rgb = RGBColor(5, 150, 105)

    r_h1 = p_t1.add_run(f"Official Onboarding & Contract Folder ({tenant_name})\n")
    r_h1.font.size = Pt(12)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(15, 23, 42)

    r_d1 = p_t1.add_run(
        "Scan this QR code with your smartphone camera to open your private Google Drive folder containing your "
        "Signed Rental Contract (Mietvertrag), City Registration Certificate (Anmeldung § 19 BMG), Keys Receipt, "
        "Handover Protocol, Inventory List, and House Rules.\n🔗 Direct Link: "
    )
    r_d1.font.size = Pt(9.5)
    add_hyperlink(p_t1, drive_folder_url, drive_folder_url, color_hex="059669", bold=True, font_size=9.0)

    # Row 1: Room KE1 Mobile Resident Hub (Linktree)
    c_img2 = qr_tbl.cell(1, 0)
    c_txt2 = qr_tbl.cell(1, 1)
    set_cell_shading(c_img2, "F0F9FF")
    set_cell_shading(c_txt2, "F0F9FF")
    set_cell_margins(c_img2, 140, 140, 140, 140)
    set_cell_margins(c_txt2, 140, 140, 180, 180)
    c_img2.width = Inches(1.7)
    c_txt2.width = Inches(5.4)

    p_i2 = c_img2.paragraphs[0]
    p_i2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_i2.add_run().add_picture(qr_lt_path, width=Inches(1.35))

    p_t2 = c_txt2.paragraphs[0]
    r_tag2 = p_t2.add_run(f"ROOM {room_code} MOBILE RESIDENT HUB & SERVICE DESK\n")
    r_tag2.font.size = Pt(8.5)
    r_tag2.font.bold = True
    r_tag2.font.color.rgb = RGBColor(2, 132, 199)

    r_h2 = p_t2.add_run(f"Homelyo Room {room_code} Digital Concierge & Support Hub\n")
    r_h2.font.size = Pt(12)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(15, 23, 42)

    r_d2 = p_t2.add_run(
        "Scan to access your mobile room hub, S+U-Bahn transit maps, neighborhood guides, and building resources.\n"
        "🛠️ Service & Maintenance Requests: For any service or repair requests, please submit an enquiry form directly "
        "from your Linktree Resident Hub below.\n🔗 Room Portal: "
    )
    r_d2.font.size = Pt(9.5)
    add_hyperlink(p_t2, linktree_url, linktree_url, color_hex="0284C7", bold=True, font_size=9.0)
    p_t2.add_run("   |   🏢 Master Hub: ").font.size = Pt(9.0)
    add_hyperlink(p_t2, master_hub_url, master_hub_url, color_hex="0284C7", bold=True, font_size=9.0)

    doc.add_paragraph()

    # Section 2: Direct Verified Document Links Table
    p_s2 = doc.add_paragraph()
    r_s2 = p_s2.add_run("2. DIRECT QUICK-LINKS TO YOUR VERIFIED DOCUMENTS")
    r_s2.font.size = Pt(11)
    r_s2.font.bold = True
    r_s2.font.color.rgb = RGBColor(15, 23, 42)

    links_tbl = doc.add_table(rows=5, cols=2)
    links_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    doc_rows = [
        ("📄 Signed Rental Contract (Mietvertrag)", contract_pdf_url),
        ("🏛️ Official City Registration (Anmeldung § 19 BMG)", anmeldung_pdf_url),
        ("🔑 Signed Keys Handover Receipt (4 Keys)", keys_pdf_url),
        ("🛠️ 24/7 Service & Maintenance Request Form (Sara AI)", service_url),
        ("📂 Complete Personal Document Folder (My_docs)", drive_folder_url),
    ]
    for idx, (label, url_val) in enumerate(doc_rows):
        row_bg = "F8FAFC" if idx % 2 == 0 else "FFFFFF"
        c0 = links_tbl.cell(idx, 0)
        c1 = links_tbl.cell(idx, 1)
        set_cell_shading(c0, row_bg)
        set_cell_shading(c1, row_bg)
        set_cell_margins(c0, 80, 80, 120, 120)
        set_cell_margins(c1, 80, 80, 120, 120)
        c0.width = Inches(3.2)
        c1.width = Inches(3.9)
        r0 = c0.paragraphs[0].add_run(label)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        add_hyperlink(c1.paragraphs[0], url_val, "Open Verified Document / Link →", color_hex="0284C7", bold=True, font_size=9.5)

    doc.add_paragraph()

    # Section 3: Essential Residence Specifications (Clean Full-Width Table)
    p_s3 = doc.add_paragraph()
    r_s3 = p_s3.add_run("3. ESSENTIAL APARTMENT & LIVING SPECIFICATIONS (WE-11)")
    r_s3.font.size = Pt(11)
    r_s3.font.bold = True
    r_s3.font.color.rgb = RGBColor(15, 23, 42)

    specs_tbl = doc.add_table(rows=4, cols=2)
    specs_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    spec_items = [
        ("📶 High-Speed Mesh Wi-Fi 6", "Network (SSID): Vodafone-327C  ·  AVM FRITZ!Box 6660 Cable (Up to 1,000 Mbps)"),
        ("🌙 German Quiet Hours (Ruhezeit)", "22:00 – 07:00 daily & all day on Sundays / Public Holidays. Peaceful co-living."),
        ("🔑 Issued Key Set (4 Keys)", f"1× Main Building Door, 1× Apartment Door (WE-11), 1× Private Room {room_code}, 1× Mailbox"),
        ("🧹 Shared Hygiene & Waste", "Weekly kitchen/bath rotation; strict Berlin waste separation (Bio, Gelber Sack, Papier, Restmüll); Stoßlüften 2× daily.")
    ]
    for idx, (k, v) in enumerate(spec_items):
        row_bg = "F8FAFC" if idx % 2 == 0 else "FFFFFF"
        c0, c1 = specs_tbl.cell(idx, 0), specs_tbl.cell(idx, 1)
        set_cell_shading(c0, row_bg)
        set_cell_shading(c1, row_bg)
        set_cell_margins(c0, 90, 90, 120, 120)
        set_cell_margins(c1, 90, 90, 120, 120)
        c0.width = Inches(2.5)
        c1.width = Inches(4.6)
        rk = c0.paragraphs[0].add_run(k)
        rk.font.bold = True
        rk.font.size = Pt(9.5)
        rv = c1.paragraphs[0].add_run(v)
        rv.font.size = Pt(9.5)

    doc.add_paragraph()

    # Section 4: Emergency & Support Contacts
    p_s4 = doc.add_paragraph()
    r_s4 = p_s4.add_run("4. OFFICIAL SUPPORT, MEDICAL & NEIGHBORHOOD CONTACTS")
    r_s4.font.size = Pt(11)
    r_s4.font.bold = True
    r_s4.font.color.rgb = RGBColor(15, 23, 42)

    em_tbl = doc.add_table(rows=6, cols=3)
    em_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    em_rows = [
        ("Homelyo Resident Helpdesk", "Official Property Management & Support", "support@homelyo.eu"),
        ("Fire & Medical Emergency", "Berliner Feuerwehr / Notarzt (24/7)", "112"),
        ("Police Dispatch (Polizei)", "Polizei Berlin Emergency Line", "110"),
        ("Nearest Hospital", "St. Marien-Hospital (Gallwitzallee 123–143, 12249 Berlin)", "030 767830"),
        ("Nearest Pharmacy", "Apotheke am Lichtenrader Damm 24, 12305 Berlin", "030 7428815"),
        ("Transit & Groceries", "Rewe Supermarket (Lichtenrader Damm 72) · S-Bahn S2 Lichtenrade", "4 min walk")
    ]
    for idx, (s_name, s_desc, s_contact) in enumerate(em_rows):
        row_bg = "F8FAFC" if idx % 2 == 0 else "FFFFFF"
        for col_i, txt in enumerate([s_name, s_desc, s_contact]):
            c = em_tbl.cell(idx, col_i)
            set_cell_shading(c, row_bg)
            set_cell_margins(c, 80, 80, 100, 100)
            r = c.paragraphs[0].add_run(txt)
            r.font.size = Pt(9.0)
            if col_i in (0, 2):
                r.font.bold = True

    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_foot = p_foot.add_run("\nHomelyo Living  ·  Lichtenrader Damm 80, 12305 Berlin  ·  support@homelyo.eu  ·  www.homelyo.eu")
    r_foot.font.size = Pt(8.5)
    r_foot.font.color.rgb = RGBColor(100, 116, 139)

    doc.save(out_docx_path)
    return out_docx_path


def build_desk_qr_stand_docx(
    tenant_name,
    room_code,
    start_date,
    end_date,
    drive_folder_url,
    linktree_url,
    out_docx_path
):
    """
    Creates a single-page, high-contrast Printable Desk QR Sheet (Word .docx)
    designed specifically to print and place/paste on Umakant's desk in Room KE1.
    """
    qr_vault_path = f"/tmp/desk_qr_vault_{room_code}.png"
    qr_lt_path = f"/tmp/desk_qr_lt_{room_code}.png"
    make_qr_png(drive_folder_url, qr_vault_path, fill_color="#0F172A")
    make_qr_png(linktree_url, qr_lt_path, fill_color="#0284C7")

    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)

    # Top Luxury Header Card
    hdr_tbl = doc.add_table(rows=1, cols=1)
    hdr_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cell = hdr_tbl.cell(0, 0)
    set_cell_shading(hdr_cell, "0F172A")
    set_cell_margins(hdr_cell, top=220, bottom=220, left=240, right=240)

    p_h = hdr_cell.paragraphs[0]
    p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_h.add_run("HOMELYO LIVING  ·  BERLIN LICHTENRADE RESIDENCE\n")
    r1.font.size = Pt(9.5)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(52, 211, 153)

    r2 = p_h.add_run(f"ROOM {room_code} — RESIDENT DESK QUICK-ACCESS HUB\n")
    r2.font.size = Pt(20)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(255, 255, 255)

    r3 = p_h.add_run(f"Welcome {tenant_name}  |  Master Bedroom Suite (21 m²)  |  Lichtenrader Damm 80, 12305 Berlin")
    r3.font.size = Pt(10.5)
    r3.font.color.rgb = RGBColor(203, 213, 225)

    doc.add_paragraph()

    # Large Primary Desk QR Box: Personal Document Vault (My_docs)
    main_tbl = doc.add_table(rows=1, cols=1)
    main_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    m_cell = main_tbl.cell(0, 0)
    set_cell_shading(m_cell, "F8FAFC")
    set_cell_margins(m_cell, top=240, bottom=240, left=260, right=260)

    pm = m_cell.paragraphs[0]
    pm.alignment = WD_ALIGN_PARAGRAPH.CENTER

    rt1 = pm.add_run("🔒 PERSONAL TENANCY & ONBOARDING DOCUMENT VAULT\n")
    rt1.font.size = Pt(13)
    rt1.font.bold = True
    rt1.font.color.rgb = RGBColor(5, 150, 105)

    rt2 = pm.add_run("Scan with your phone camera to open your private Google Drive folder (My_docs):\n")
    rt2.font.size = Pt(10.5)
    rt2.font.color.rgb = RGBColor(51, 65, 85)

    p_qr_img = m_cell.add_paragraph()
    p_qr_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_qr_img.add_run().add_picture(qr_vault_path, width=Inches(2.45))

    p_qr_cap = m_cell.add_paragraph()
    p_qr_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rc1 = p_qr_cap.add_run(
        "Includes: Signed Rental Contract (Mietvertrag) · Official Anmeldung (§ 19 BMG) · "
        "Welcome Kit · Keys Receipt · Inventory & House Rules\n"
    )
    rc1.font.size = Pt(9.5)
    rc1.font.bold = True
    rc1.font.color.rgb = RGBColor(15, 23, 42)
    add_hyperlink(p_qr_cap, drive_folder_url, drive_folder_url, color_hex="059669", bold=True, font_size=9.5)

    doc.add_paragraph()

    # Secondary Row: Left = Room KE1 Mobile Hub QR | Right = Wi-Fi & Support Info
    sub_tbl = doc.add_table(rows=1, cols=2)
    sub_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_left = sub_tbl.cell(0, 0)
    c_right = sub_tbl.cell(0, 1)
    set_cell_shading(c_left, "EFF6FF")
    set_cell_shading(c_right, "F0FDF4")
    set_cell_margins(c_left, top=180, bottom=180, left=180, right=180)
    set_cell_margins(c_right, top=180, bottom=180, left=180, right=180)
    c_left.width = Inches(3.5)
    c_right.width = Inches(3.5)

    pl = c_left.paragraphs[0]
    pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rl1 = pl.add_run(f"📲 ROOM {room_code} MOBILE HUB & SERVICE\n")
    rl1.font.size = Pt(10.5)
    rl1.font.bold = True
    rl1.font.color.rgb = RGBColor(2, 132, 199)

    pl_img = c_left.add_paragraph()
    pl_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pl_img.add_run().add_picture(qr_lt_path, width=Inches(1.45))

    pl_cap = c_left.add_paragraph()
    pl_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rl2 = pl_cap.add_run("🛠️ For any service or repair requests, submit an enquiry form via your Linktree Hub:\n")
    rl2.font.size = Pt(8.5)
    add_hyperlink(pl_cap, linktree_url, linktree_url, color_hex="0284C7", bold=True, font_size=8.5)

    pr = c_right.paragraphs[0]
    rr1 = pr.add_run("📶 QUICK DESK REFERENCE\n\n")
    rr1.font.size = Pt(10.5)
    rr1.font.bold = True
    rr1.font.color.rgb = RGBColor(6, 95, 70)

    info_lines = [
        ("Wi-Fi Network (SSID): ", "Vodafone-327C (Wi-Fi 6 Mesh)\n"),
        ("Quiet Hours (Ruhezeit): ", "22:00 – 07:00 Daily\n"),
        ("Apartment Unit: ", f"2. OG · WE-11 · Room {room_code}\n"),
        ("Support Email: ", "support@homelyo.eu\n"),
        ("24/7 AI Concierge: ", "homelyo.eu/assistant\n"),
        ("Emergency (Fire / Police): ", "112  /  110")
    ]
    for k_lbl, v_lbl in info_lines:
        rk = pr.add_run(k_lbl)
        rk.font.bold = True
        rk.font.size = Pt(9.5)
        rv = pr.add_run(v_lbl)
        rv.font.size = Pt(9.5)

    p_f = doc.add_paragraph()
    p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rf = p_f.add_run("\nHomelyo Living  ·  Lichtenrader Damm 80, 12305 Berlin  ·  support@homelyo.eu")
    rf.font.size = Pt(8.5)
    rf.font.color.rgb = RGBColor(100, 116, 139)

    doc.save(out_docx_path)
    return out_docx_path


def render_executive_welcome_html(
    tenant_name,
    room_code,
    start_date,
    end_date,
    rent_amount,
    deposit_amount,
    drive_folder_url,
    contract_pdf_url,
    anmeldung_pdf_url,
    keys_pdf_url
):
    room_info = ROOM_SPECS.get(room_code, {"title": "Master Bedroom Suite", "size": 21})
    linktree_url = LINKTREE_URLS.get(room_code, "https://linktr.ee/lichtenraderdamm_80_ke_1")
    master_hub_url = "https://linktr.ee/lichtenrder_damm_80"
    service_url = "https://www.homelyo.eu/assistant/"

    qr_vault_b64 = generate_qr_data_uri(drive_folder_url, fill_color="#065F46")
    qr_lt_b64 = generate_qr_data_uri(linktree_url, fill_color="#0F172A")

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Welcome Kit — {tenant_name} ({room_code})</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');
  @page {{
    size: A4 portrait;
    margin: 12mm 14mm 12mm 14mm;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    color: #0F172A;
    background: #FFFFFF;
    line-height: 1.48;
    font-size: 9.5pt;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }}
  .page {{
    page-break-after: always;
    position: relative;
  }}
  .page:last-child {{
    page-break-after: avoid;
  }}
  .top-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding-bottom: 10px;
    border-bottom: 2px solid #0F172A;
    margin-bottom: 14px;
  }}
  .brand-title {{
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 16pt;
    font-weight: 800;
    letter-spacing: -0.4px;
    color: #0F172A;
  }}
  .brand-title span {{ color: #059669; }}
  .brand-tag {{
    font-size: 8pt;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    color: #64748B;
  }}
  .pill-badge {{
    background: #0F172A;
    color: #34D399;
    font-size: 8pt;
    font-weight: 700;
    padding: 5px 12px;
    border-radius: 99px;
    letter-spacing: 0.3px;
  }}
  .hero-banner {{
    background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
    color: #FFFFFF;
    border-radius: 14px;
    padding: 18px 22px;
    margin-bottom: 16px;
  }}
  .hero-kicker {{
    color: #34D399;
    font-size: 8pt;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 4px;
  }}
  .hero-name {{
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 17pt;
    font-weight: 800;
    margin-bottom: 6px;
  }}
  .hero-desc {{
    font-size: 9.2pt;
    color: #CBD5E1;
    margin-bottom: 12px;
  }}
  .spec-strip {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 10px;
    padding-top: 12px;
    border-top: 1px solid rgba(255,255,255,0.14);
  }}
  .spec-box .lbl {{
    font-size: 7pt;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    color: #94A3B8;
    font-weight: 600;
  }}
  .spec-box .val {{
    font-size: 9pt;
    font-weight: 700;
    color: #F8FAFC;
    margin-top: 2px;
  }}
  .sec-head {{
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 10.5pt;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    color: #0F172A;
    margin-bottom: 10px;
    padding-left: 8px;
    border-left: 3.5px solid #059669;
  }}
  /* Wide Horizontal QR Cards — No narrow vertical boxes! */
  .qr-row-card {{
    display: flex;
    align-items: center;
    gap: 18px;
    border-radius: 12px;
    padding: 14px 18px;
    margin-bottom: 12px;
    border: 1.5px solid #CBD5E1;
  }}
  .qr-row-card.vault {{
    background: #F0FDF4;
    border-color: #86EFAC;
  }}
  .qr-row-card.hub {{
    background: #F0F9FF;
    border-color: #93C5FD;
  }}
  .qr-frame {{
    background: #FFFFFF;
    padding: 8px;
    border-radius: 10px;
    border: 1px solid #E2E8F0;
    flex-shrink: 0;
  }}
  .qr-frame img {{
    width: 112px;
    height: 112px;
    display: block;
  }}
  .qr-body {{
    flex: 1;
  }}
  .qr-pill {{
    display: inline-block;
    font-size: 7.2pt;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 0.6px;
    padding: 2px 8px;
    border-radius: 99px;
    margin-bottom: 4px;
  }}
  .qr-pill.vault {{ background: #059669; color: #FFFFFF; }}
  .qr-pill.hub {{ background: #0284C7; color: #FFFFFF; }}
  .qr-heading {{
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 11.5pt;
    font-weight: 800;
    color: #0F172A;
    margin-bottom: 4px;
  }}
  .qr-text {{
    font-size: 8.6pt;
    color: #334155;
    line-height: 1.42;
    margin-bottom: 6px;
  }}
  .qr-url {{
    font-size: 8.3pt;
    font-weight: 700;
    color: #0284C7;
    text-decoration: underline;
    word-break: break-all;
  }}
  .qr-url.vault {{ color: #059669; }}
  .service-banner {{
    margin-top: 6px;
    padding: 6px 10px;
    background: #FFFFFF;
    border-left: 3px solid #0284C7;
    border-radius: 6px;
    font-size: 8pt;
    color: #1E3A8A;
  }}
  .doc-table {{
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 14px;
    font-size: 8.6pt;
  }}
  .doc-table th {{
    background: #0F172A;
    color: #FFFFFF;
    text-align: left;
    padding: 6px 10px;
    font-size: 8pt;
    text-transform: uppercase;
    letter-spacing: 0.4px;
  }}
  .doc-table td {{
    padding: 7px 10px;
    border-bottom: 1px solid #E2E8F0;
    color: #1E293B;
  }}
  .doc-table tr:nth-child(even) {{ background: #F8FAFC; }}
  .doc-table a {{
    color: #0284C7;
    font-weight: 700;
    text-decoration: none;
  }}
  .about-box {{
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 10px;
  }}
  .about-box p {{
    font-size: 8.6pt;
    color: #334155;
    margin-bottom: 4px;
  }}
  .about-box p.de {{
    font-size: 8.2pt;
    color: #64748B;
    font-style: italic;
    margin-bottom: 0;
  }}
  .footer {{
    text-align: center;
    font-size: 7.8pt;
    color: #64748B;
    padding-top: 8px;
    border-top: 1px solid #E2E8F0;
    margin-top: 8px;
  }}
</style>
</head>
<body>

<!-- PAGE 1 -->
<div class="page">
  <div class="top-header">
    <div>
      <div class="brand-title">HOMELYO<span>.EU</span></div>
      <div class="brand-tag">Berlin Lichtenrade Boutique Co-Living Residence</div>
    </div>
    <div class="pill-badge">✓ OFFICIAL RESIDENT DOSSIER · ROOM {room_code}</div>
  </div>

  <div class="hero-banner">
    <div class="hero-kicker">Welcome Home to Berlin</div>
    <div class="hero-name">{tenant_name} — Room {room_code}</div>
    <div class="hero-desc">
      Your private suite and official tenancy documentation at Lichtenrader Damm 80 (2. OG, WE-11) are ready for your arrival.
    </div>
    <div class="spec-strip">
      <div class="spec-box">
        <div class="lbl">Assigned Suite</div>
        <div class="val">Room {room_code} ({room_info['size']} m²)</div>
      </div>
      <div class="spec-box">
        <div class="lbl">Tenancy Period</div>
        <div class="val">{start_date} – {end_date}</div>
      </div>
      <div class="spec-box">
        <div class="lbl">All-In Warm Rent</div>
        <div class="val">€{rent_amount:,.2f} / month</div>
      </div>
      <div class="spec-box">
        <div class="lbl">Property Management</div>
        <div class="val">Homelyo Living</div>
      </div>
    </div>
  </div>

  <div class="sec-head">1. Digital Access & Personal Cloud Vault (Scan or Click)</div>

  <!-- Horizontal Card 1: Personal Document Vault -->
  <div class="qr-row-card vault">
    <div class="qr-frame">
      <a href="{drive_folder_url}" target="_blank">
        <img src="{qr_vault_b64}" alt="Personal Vault QR">
      </a>
    </div>
    <div class="qr-body">
      <div class="qr-pill vault">Private Tenant Vault · Personal Access Only</div>
      <div class="qr-heading">Your Dedicated Google Drive Document Folder (My_docs)</div>
      <div class="qr-text">
        Scan this QR code with your smartphone camera to open your private folder containing all 8 official onboarding documents: Signed Rental Contract (Mietvertrag), City Registration (Anmeldung § 19 BMG), Keys Receipt, Inventory List & House Rules.
      </div>
      <a href="{drive_folder_url}" target="_blank" class="qr-url vault">{drive_folder_url}</a>
    </div>
  </div>

  <!-- Horizontal Card 2: Room KE1 Linktree Mobile Hub -->
  <div class="qr-row-card hub">
    <div class="qr-frame">
      <a href="{linktree_url}" target="_blank">
        <img src="{qr_lt_b64}" alt="Room Mobile Hub QR">
      </a>
    </div>
    <div class="qr-body">
      <div class="qr-pill hub">Room {room_code} Mobile Hub & Concierge</div>
      <div class="qr-heading">Room {room_code} Digital Resident Portal & Support Hub</div>
      <div class="qr-text">
        Scan to open your mobile room portal with 1-tap access to house rules, Berlin S+U-Bahn transit map, neighborhood street guide, and the <a href="{master_hub_url}" target="_blank" style="color:#0284C7;font-weight:600;">Building Master Hub</a>.
      </div>
      <a href="{linktree_url}" target="_blank" class="qr-url">{linktree_url}</a>
      <div class="service-banner">
        🛠️ <strong>Service & Maintenance Requests:</strong> For any service requests, please send an enquiry form directly from your Linktree Resident Hub above (<a href="{service_url}" target="_blank" style="color:#0284C7;font-weight:700;">homelyo.eu/assistant</a>).
      </div>
    </div>
  </div>

  <div class="sec-head">2. Direct Links to Verified Tenancy Documents</div>
  <table class="doc-table">
    <thead>
      <tr>
        <th>Document / Resource</th>
        <th>Description</th>
        <th>Direct Cloud Link</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td><strong>📄 Rental Contract (Mietvertrag)</strong></td>
        <td>Bilingual German/English Tenancy Agreement</td>
        <td><a href="{contract_pdf_url}" target="_blank">Open Contract PDF →</a></td>
      </tr>
      <tr>
        <td><strong>🏛️ City Registration (Anmeldung)</strong></td>
        <td>Official Wohnungsgeberbestätigung (§ 19 BMG)</td>
        <td><a href="{anmeldung_pdf_url}" target="_blank">Open Certificate PDF →</a></td>
      </tr>
      <tr>
        <td><strong>🔑 Key Handover Receipt</strong></td>
        <td>Official Receipt for 4 Issued Security Keys</td>
        <td><a href="{keys_pdf_url}" target="_blank">Open Keys Protocol →</a></td>
      </tr>
      <tr>
        <td><strong>📂 Complete Personal Vault (My_docs)</strong></td>
        <td>All 8 Onboarding PDFs & Handover Checklists</td>
        <td><a href="{drive_folder_url}" target="_blank">Open Google Drive Vault →</a></td>
      </tr>
    </tbody>
  </table>

  <div class="about-box">
    <p><strong>🏡 About Lichtenrader Damm 80:</strong> Located in the peaceful, leafy district of Berlin-Lichtenrade with supermarkets (Rewe 4 min walk), bakeries, pharmacies, and direct S-Bahn Line S2 connection to central Berlin (Potsdamer Platz / Friedrichstraße).</p>
    <p class="de">Lichtenrader Damm 80 liegt im ruhigen Berliner Ortsteil Lichtenrade mit direkter S-Bahn-Anbindung (S2) ins Stadtzentrum.</p>
  </div>

  <div class="footer">
    Homelyo Living · Lichtenrader Damm 80, 12305 Berlin · support@homelyo.eu · Page 1 of 2
  </div>
</div>

<!-- PAGE 2 -->
<div class="page">
  <div class="top-header">
    <div>
      <div class="brand-title">HOMELYO<span>.EU</span></div>
      <div class="brand-tag">Residence Specifications, House Rules & Emergency Directory</div>
    </div>
    <div class="pill-badge">ROOM {room_code} · {tenant_name.upper()}</div>
  </div>

  <div class="sec-head">3. Apartment Living Guide & Key Specifications (WE-11)</div>
  <table class="doc-table" style="margin-bottom: 16px;">
    <thead>
      <tr>
        <th style="width: 28%;">Topic / Facility</th>
        <th>Official Standard & Instructions</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td><strong>📶 High-Speed Wi-Fi 6</strong></td>
        <td><strong>Network Name (SSID):</strong> Vodafone-327C &nbsp;|&nbsp; <strong>Router:</strong> AVM FRITZ!Box 6660 Cable (Up to 1,000 Mbps Mesh coverage across all rooms).</td>
      </tr>
      <tr>
        <td><strong>🔑 Key Inventory (4 Keys)</strong></td>
        <td>You receive <strong>4 keys</strong>: (1) Main Building Entrance, (2) Apartment Door WE-11, (3) Private Room {room_code}, (4) Mailbox. Please keep keys safe; lock replacement is costly.</td>
      </tr>
      <tr>
        <td><strong>🌙 Quiet Hours (Ruhezeit)</strong></td>
        <td><strong>22:00 – 07:00 daily</strong> and all day on Sundays & German public holidays. Please keep noise at room volume so all residents can rest and study peacefully.</td>
      </tr>
      <tr>
        <td><strong>🧹 Shared Hygiene & Waste</strong></td>
        <td>Kitchen and bathroom cleaning rotates weekly among the 4 residents. Please separate waste strictly into <strong>Bio</strong>, <strong>Gelber Sack</strong> (packaging), <strong>Papier</strong> (blue), and <strong>Restmüll</strong> (grey).</td>
      </tr>
      <tr>
        <td><strong>🪟 Daily Ventilation</strong></td>
        <td>Perform <strong>Stoßlüften</strong> (open windows wide for 5–10 minutes twice daily with heating turned down) to maintain fresh air and prevent condensation.</td>
      </tr>
      <tr>
        <td><strong>🛠️ Service & Repairs</strong></td>
        <td>Report any maintenance issues immediately via the <strong>Service & Maintenance Request Form</strong> on your Linktree Hub (<a href="{service_url}" target="_blank">homelyo.eu/assistant</a>) or email <a href="mailto:support@homelyo.eu">support@homelyo.eu</a>.</td>
      </tr>
    </tbody>
  </table>

  <div class="sec-head">4. Official Support, Medical & Neighborhood Directory</div>
  <table class="doc-table" style="margin-bottom: 16px;">
    <thead>
      <tr>
        <th style="width: 30%;">Service / Institution</th>
        <th>Address & Details</th>
        <th style="width: 24%;">Contact / Hotline</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td><strong>Homelyo Resident Support</strong></td>
        <td>Official Property Management & Tenancy Helpdesk</td>
        <td><a href="mailto:support@homelyo.eu">support@homelyo.eu</a></td>
      </tr>
      <tr>
        <td><strong>24/7 AI Concierge (Sara)</strong></td>
        <td>Instant Maintenance Tickets & Resident Assistant</td>
        <td><a href="{service_url}" target="_blank">homelyo.eu/assistant</a></td>
      </tr>
      <tr>
        <td><strong>Fire & Medical Emergency</strong></td>
        <td>Berliner Feuerwehr / Notarzt (24/7 Dispatch)</td>
        <td><strong style="color:#DC2626;">112 ☎️</strong></td>
      </tr>
      <tr>
        <td><strong>Police (Polizei Berlin)</strong></td>
        <td>Emergency Police Dispatch (Abschnitt 46)</td>
        <td><strong style="color:#0284C7;">110 🚓</strong></td>
      </tr>
      <tr>
        <td><strong>Nearest Hospitals</strong></td>
        <td>St. Marien-Hospital (Gallwitzallee 123–143, 12249 Berlin)<br>Vivantes Wenckebach-Klinikum (Wenckebachstr. 23, 12099 Berlin)</td>
        <td>030 767830 🏥<br>030 130190 🏥</td>
      </tr>
      <tr>
        <td><strong>Nearest Pharmacies</strong></td>
        <td>Apotheke am Lichtenrader Damm (Lichtenrader Damm 49)<br>Grüne Apotheke Lichtenrade (Lichtenrader Damm 24)</td>
        <td>030 743048690 💊<br>030 7428815 💊</td>
      </tr>
      <tr>
        <td><strong>Groceries & S-Bahn</strong></td>
        <td>Rewe Supermarket (Lichtenrader Damm 72, 4 min walk)<br>S-Bahnhof Lichtenrade (Line S2 direct to city center)</td>
        <td><a href="https://maps.google.com/?q=Lichtenrader+Damm+80+12305+Berlin" target="_blank">Open Google Maps →</a></td>
      </tr>
    </tbody>
  </table>

  <div class="about-box" style="background:#ECFDF5; border-color:#86EFAC;">
    <p style="color:#065F46; margin-bottom:2px;"><strong>✨ Pre-Arrival Preparation Guarantee (Move-In Checklist Completed):</strong></p>
    <p style="color:#047857; margin-bottom:0;">
      ☑ Private room {room_code} professionally deep-cleaned &nbsp;·&nbsp; ☑ Fresh sanitized pillows, mattress protector & bed linens &nbsp;·&nbsp; ☑ High-speed Wi-Fi 6 verified &nbsp;·&nbsp; ☑ Handover protocol & 4-key set prepared for your arrival on <strong>{start_date}</strong>.
    </p>
  </div>

  <div style="text-align:center; margin-top:14px; font-family:'Plus Jakarta Sans',sans-serif; font-weight:800; font-size:10.5pt; color:#0F172A;">
    Genießen Sie Ihren Aufenthalt! · Enjoy Your Stay at Homelyo!
  </div>

  <div class="footer" style="margin-top:12px;">
    Homelyo Living · Lichtenrader Damm 80, 12305 Berlin · support@homelyo.eu · Page 2 of 2
  </div>
</div>

</body>
</html>"""


def render_desk_qr_stand_html(tenant_name, room_code, drive_folder_url, linktree_url):
    qr_vault_b64 = generate_qr_data_uri(drive_folder_url, fill_color="#0F172A")
    qr_lt_b64 = generate_qr_data_uri(linktree_url, fill_color="#0284C7")
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Printable Desk QR Sheet — Room {room_code}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');
  @page {{ size: A4 portrait; margin: 14mm; }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Inter', sans-serif;
    color: #0F172A;
    background: #FFFFFF;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }}
  .stand-frame {{
    border: 3px solid #0F172A;
    border-radius: 20px;
    padding: 24px;
    background: #FFFFFF;
  }}
  .banner {{
    background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
    color: #FFFFFF;
    border-radius: 14px;
    padding: 20px 24px;
    text-align: center;
    margin-bottom: 20px;
  }}
  .banner .kicker {{
    color: #34D399;
    font-size: 9pt;
    font-weight: 800;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    margin-bottom: 4px;
  }}
  .banner h1 {{
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 20pt;
    font-weight: 800;
    margin-bottom: 4px;
  }}
  .banner p {{
    font-size: 10pt;
    color: #CBD5E1;
  }}
  .primary-vault {{
    background: #F0FDF4;
    border: 2px solid #10B981;
    border-radius: 16px;
    padding: 22px;
    text-align: center;
    margin-bottom: 20px;
  }}
  .primary-vault .badge {{
    display: inline-block;
    background: #059669;
    color: #FFFFFF;
    font-size: 8.5pt;
    font-weight: 800;
    padding: 4px 14px;
    border-radius: 99px;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    margin-bottom: 8px;
  }}
  .primary-vault h2 {{
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 15pt;
    font-weight: 800;
    color: #0F172A;
    margin-bottom: 6px;
  }}
  .primary-vault p {{
    font-size: 9.5pt;
    color: #334155;
    margin-bottom: 14px;
  }}
  .qr-big-box {{
    display: inline-block;
    background: #FFFFFF;
    padding: 14px;
    border-radius: 16px;
    border: 2px solid #A7F3D0;
    box-shadow: 0 4px 12px rgba(0,0,0,0.06);
    margin-bottom: 12px;
  }}
  .qr-big-box img {{
    width: 210px;
    height: 210px;
    display: block;
  }}
  .url-text {{
    font-size: 8.8pt;
    font-weight: 700;
    color: #059669;
    word-break: break-all;
  }}
  .bottom-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
  }}
  .sub-card {{
    border-radius: 14px;
    padding: 16px;
    border: 1.5px solid #CBD5E1;
    background: #F8FAFC;
  }}
  .sub-card.blue {{
    background: #EFF6FF;
    border-color: #93C5FD;
    text-align: center;
  }}
  .sub-card h3 {{
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 11pt;
    font-weight: 800;
    margin-bottom: 8px;
    color: #0F172A;
  }}
  .qr-sm-box {{
    display: inline-block;
    background: #FFFFFF;
    padding: 8px;
    border-radius: 10px;
    border: 1px solid #BFDBFE;
    margin-bottom: 8px;
  }}
  .qr-sm-box img {{
    width: 115px;
    height: 115px;
    display: block;
  }}
  .info-list {{
    list-style: none;
    font-size: 9pt;
    line-height: 1.65;
  }}
  .info-list li {{
    border-bottom: 1px dashed #CBD5E1;
    padding: 3px 0;
  }}
  .info-list li:last-child {{ border-bottom: none; }}
  .cut-footer {{
    text-align: center;
    margin-top: 14px;
    font-size: 8.5pt;
    color: #64748B;
    font-weight: 600;
  }}
</style>
</head>
<body>
  <div class="stand-frame">
    <div class="banner">
      <div class="kicker">Homelyo Living · Lichtenrader Damm 80, 12305 Berlin</div>
      <h1>ROOM {room_code} — RESIDENT DESK HUB</h1>
      <p>Welcome {tenant_name} · Master Bedroom Suite (21 m²) · WE-11 (2. OG)</p>
    </div>

    <div class="primary-vault">
      <div class="badge">🔒 Scan for Your Personal Document Vault</div>
      <h2>Official Tenancy & Onboarding Folder (My_docs)</h2>
      <p>Point your smartphone camera at this QR code to open your private Google Drive folder<br>(Signed Rental Contract · Anmeldung § 19 BMG · Welcome Kit · Keys Receipt · House Rules):</p>
      <div class="qr-big-box">
        <img src="{qr_vault_b64}" alt="Personal Document Vault QR">
      </div>
      <div class="url-text">{drive_folder_url}</div>
    </div>

    <div class="bottom-grid">
      <div class="sub-card blue">
        <h3>📲 Room {room_code} Mobile Hub & Service</h3>
        <div class="qr-sm-box">
          <img src="{qr_lt_b64}" alt="Linktree Hub QR">
        </div>
        <div style="font-size:8pt; color:#1E3A8A; line-height:1.35;">
          🛠️ <strong>Service & Repairs:</strong> Submit any maintenance enquiry directly via your Linktree Hub:<br>
          <strong>{linktree_url.replace('https://', '')}</strong>
        </div>
      </div>

      <div class="sub-card">
        <h3>📶 Desk Quick-Reference</h3>
        <ul class="info-list">
          <li><strong>Wi-Fi Network:</strong> Vodafone-327C (Wi-Fi 6)</li>
          <li><strong>Quiet Hours:</strong> 22:00 – 07:00 Daily</li>
          <li><strong>Apartment Unit:</strong> 2. OG · WE-11 · Room {room_code}</li>
          <li><strong>Support Email:</strong> support@homelyo.eu</li>
          <li><strong>24/7 AI Concierge:</strong> homelyo.eu/assistant</li>
          <li><strong>Emergency:</strong> Fire/Medical 112 · Police 110</li>
        </ul>
      </div>
    </div>

    <div class="cut-footer">
      🖨️ Official Room {room_code} Desk Reference Sheet · Homelyo Living · www.homelyo.eu
    </div>
  </div>
</body>
</html>"""


def html_to_pdf(html_str, out_pdf_path):
    tmp_html = f"/tmp/render_{abs(hash(out_pdf_path))}.html"
    with open(tmp_html, "w", encoding="utf-8") as f:
        f.write(html_str)
    chrome_cmd = [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "--headless",
        "--disable-gpu",
        "--no-pdf-header-footer",
        f"--print-to-pdf={out_pdf_path}",
        tmp_html
    ]
    subprocess.run(chrome_cmd, check=True)
    return out_pdf_path


def upload_or_update_drive_file(token, folder_id, filename, local_path, mime_type, convert_to_gdoc=False):
    # Check if file with same name and target mimeType exists in folder_id
    target_mime = "application/vnd.google-apps.document" if convert_to_gdoc else mime_type
    q = f"'{folder_id}' in parents and name = '{filename}' and mimeType = '{target_mime}' and trashed = false"
    q_url = f"https://www.googleapis.com/drive/v3/files?q={urllib.parse.quote(q)}&fields=files(id,name)"
    req = urllib.request.Request(q_url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req) as r:
        existing = json.loads(r.read().decode()).get("files", [])

    with open(local_path, "rb") as f:
        file_bytes = f.read()

    if existing and not convert_to_gdoc:
        fid = existing[0]["id"]
        up_url = f"https://www.googleapis.com/upload/drive/v3/files/{fid}?uploadType=media"
        up_req = urllib.request.Request(
            up_url,
            data=file_bytes,
            headers={"Authorization": f"Bearer {token}", "Content-Type": mime_type},
            method="PATCH"
        )
        with urllib.request.urlopen(up_req) as resp:
            res = json.loads(resp.read().decode())
            print(f"  ✓ Updated existing Drive file {filename} (id={fid})")
            return fid
    else:
        # Delete old Google Doc if replacing a converted Google Doc
        for old_f in existing:
            del_req = urllib.request.Request(
                f"https://www.googleapis.com/drive/v3/files/{old_f['id']}",
                headers={"Authorization": f"Bearer {token}"},
                method="DELETE"
            )
            try:
                urllib.request.urlopen(del_req)
            except Exception:
                pass

        boundary = "===HOMELYO_UPLOAD_BOUNDARY==="
        meta = {"name": filename, "parents": [folder_id]}
        if convert_to_gdoc:
            meta["mimeType"] = "application/vnd.google-apps.document"

        body = (
            f"--{boundary}\r\n"
            "Content-Type: application/json; charset=UTF-8\r\n\r\n"
            f"{json.dumps(meta)}\r\n"
            f"--{boundary}\r\n"
            f"Content-Type: {mime_type}\r\n\r\n"
        ).encode("utf-8") + file_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")

        create_req = urllib.request.Request(
            "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink",
            data=body,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": f"multipart/related; boundary={boundary}"
            },
            method="POST"
        )
        with urllib.request.urlopen(create_req) as resp:
            res = json.loads(resp.read().decode())
            print(f"  ✓ Uploaded {filename} to folder {folder_id} -> id={res.get('id')}")
            return res.get("id")


if __name__ == "__main__":
    import urllib.parse
    token = get_gcloud_token()

    umakant_main_folder = "10wcLSRWAzah9EGwCRYhgwADhWVIF_Gzn"
    umakant_mydocs_folder = "1x6WqAM3ap1xi0EGKNWRcrafycA0y-wiU"
    mydocs_url = "https://drive.google.com/drive/folders/1x6WqAM3ap1xi0EGKNWRcrafycA0y-wiU"
    ke1_linktree_url = "https://linktr.ee/lichtenraderdamm_80_ke_1"

    contract_url = "https://drive.google.com/file/d/1_VZQHkXwq0rAk27GiND1mD5fd8ZN1gDU/view"
    anmeldung_url = "https://drive.google.com/file/d/1R-fiY2Wb6ZI7rKb0HSxKVmi4gfSzUGcs/view"
    keys_url = "https://drive.google.com/file/d/1U-Ya3CQ6jCHY_w8mjd0keoTmADkk1ba0/view"

    # 1. Build Redesigned Executive Welcome Kit (DOCX + PDF)
    wk_docx = "/tmp/Welcome_Kit_KE1_Umakant.docx"
    wk_pdf = "/tmp/Welcome_Kit_KE1_Umakant.pdf"

    build_welcome_kit_docx(
        tenant_name="Umakant Yadav",
        room_code="KE1",
        start_date="01 November 2026",
        end_date="30 April 2027",
        rent_amount=609.00,
        deposit_amount=1000.00,
        drive_folder_url=mydocs_url,
        contract_pdf_url=contract_url,
        anmeldung_pdf_url=anmeldung_url,
        keys_pdf_url=keys_url,
        out_docx_path=wk_docx
    )

    wk_html = render_executive_welcome_html(
        tenant_name="Umakant Yadav",
        room_code="KE1",
        start_date="01 November 2026",
        end_date="30 April 2027",
        rent_amount=609.00,
        deposit_amount=1000.00,
        drive_folder_url=mydocs_url,
        contract_pdf_url=contract_url,
        anmeldung_pdf_url=anmeldung_url,
        keys_pdf_url=keys_url
    )
    html_to_pdf(wk_html, wk_pdf)

    # 2. Build Printable Desk QR Stand (Word DOCX + Google Doc + PDF)
    desk_docx = "/tmp/Printable_Desk_QR_Sheet_KE1_Umakant.docx"
    desk_pdf = "/tmp/Printable_Desk_QR_Sheet_KE1_Umakant.pdf"

    build_desk_qr_stand_docx(
        tenant_name="Umakant Yadav",
        room_code="KE1",
        start_date="01 November 2026",
        end_date="30 April 2027",
        drive_folder_url=mydocs_url,
        linktree_url=ke1_linktree_url,
        out_docx_path=desk_docx
    )

    desk_html = render_desk_qr_stand_html(
        tenant_name="Umakant Yadav",
        room_code="KE1",
        drive_folder_url=mydocs_url,
        linktree_url=ke1_linktree_url
    )
    html_to_pdf(desk_html, desk_pdf)

    print("🚀 Uploading redesigned Welcome Kit and Printable Desk QR Word Documents to Umakant's folders...")
    docx_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    # Update Welcome_Kit_KE1_Umakant.pdf in My_docs (preserving file ID 10aVMgMfbv8a9Yjv687dy2U3D-SDODorU!)
    upload_or_update_drive_file(token, umakant_mydocs_folder, "Welcome_Kit_KE1_Umakant.pdf", wk_pdf, "application/pdf", convert_to_gdoc=False)
    # Also upload editable Word .docx & Google Doc of Welcome Kit in My_docs
    upload_or_update_drive_file(token, umakant_mydocs_folder, "Welcome_Kit_KE1_Umakant.docx", wk_docx, docx_mime, convert_to_gdoc=False)

    # Upload Printable Desk QR Sheet (Word .docx, Google Doc & PDF) in Umakant's Main Folder AND My_docs Folder
    for target_f in [umakant_main_folder, umakant_mydocs_folder]:
        upload_or_update_drive_file(token, target_f, "Printable_Desk_QR_Sheet_KE1_Umakant.docx", desk_docx, docx_mime, convert_to_gdoc=False)
        upload_or_update_drive_file(token, target_f, "Printable_Desk_QR_Sheet_KE1_Umakant (Google Doc)", desk_docx, docx_mime, convert_to_gdoc=True)
        upload_or_update_drive_file(token, target_f, "Printable_Desk_QR_Sheet_KE1_Umakant.pdf", desk_pdf, "application/pdf", convert_to_gdoc=False)

    print("🎉 Successfully generated and uploaded all redesigned Welcome Kits and Printable Desk QR Word documents!")

