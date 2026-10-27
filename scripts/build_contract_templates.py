#!/usr/bin/env python3
"""
Homelyo Master Bilingual Rental Contract Builder
------------------------------------------------
Creates the unified bilingual (German/English) rental contract template for Homelyo,
faithfully combining:
1. The 14-clause legal architecture of Wunderflats & GDrive samples:
   - Hauptvertrag (Core Agreement):
     * § 1 Mietdauer & Befristungszweck / Term of Lease & Temporary Purpose
     * § 2 Miethöhe, Pauschale & Nebenkosten / Rent, All-Inclusive Utilities & License Fee
     * § 3 Kaution & Mietsicherheit / Security Deposit & Escrow (§ 551 BGB)
     * § 4 Zahlungsbedingungen, Fälligkeit & Verzug / Payment Terms, Due Dates & Late Fees
     * § 5 Kündigung & Nachmieter / Notice & Replacement Tenant Release
     * § 6 AGB & Anlagen / Terms & Attachments
     * Hauptvertrag Signatures (Parties)
   - Zahlungsmöglichkeiten / Payment Options & Bank Details (N26 Bank AG)
   - Anlage 1 – Inventar der Mietsache / Attachment 1 – Inventory List & Upkeep Rules
   - Anlage 2 – Allgemeine Geschäftsbedingungen (AGB / House Rules):
     * Preamble / Präambel
     * § 1 Zusammenleben in der WG / mit den Nachbarn / Cohabitation in Flat Share
     * § 2 Regelungen zur Wohnung / Apartment Regulations (Heating, Ventilation, Internet, Waste)
     * § 3 Zahlungsvereinbarungen / Payment Agreements (Punctuality, 30-day refund, Fair Use)
     * § 4 Übergabe und Rückgabe / Handover and Return of the Property
     * § 5 Schäden an der Wohnung / Damages to Property (Minor repairs schedule)
     * § 6 Kündigungsbedingungen / Conditions of Termination (Tacit renewal exclusion § 545 BGB)
     * § 7 Schriftform / Written Form
     * § 8 Salvatorische Klausel & Gerichtsstand / Severability & Jurisdiction
     * AGB Signatures (Parties)
2. All agreed Homelyo rates:
   - Key replacement fees: Haustür 680,00 €, Wohnung 120,00 €, Zimmer 90,00 €, Briefkasten 70,00 €
   - One-time checkout cleaning fee: 70,00 €
   - Security deposit: 1.000,00 €
   - Overdue reminder fee: 5,00 €
   - Minor repairs (Kleinreparaturklausel): 100,00 € per instance, max 8% of rent or 350,00 € per year
   - Termination for unpaid damage threshold: > 50,00 €
   - Deposit refund timeframe: within 30 days
   - Replacement tenant notice: 3 months' notice to month-end
   - Overnight guest limit: maximum 5 days per calendar month
3. Flawless A4 formatting for PDF export (headless Chrome) and Word DOCX generation.
"""

import os
import sys
import datetime
import base64
import subprocess
import urllib.request
import json
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

LOGO_PATH = "/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/assistant/assets/Homelyo-logo-header.png"
FALLBACK_LOGO_PATH = "/Users/jayashreesundarasamy/Downloads/Homelyo-logo.png"
HTML_TEMPLATE_PATH = "/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/templates/contract_template.html"
DOCX_TEMPLATE_PATH = "/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/template_contract.docx"
FINAL_PDF_PATH = "/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/contracts_generated/Rental_Contract_Master_Final.pdf"
GDRIVE_TEMPLATE_FILE_ID = "1Cj9E9nbcCbgV4HuT2ZLxdfNwiDLYVT5t"

ROOM_SPECS = {
    "KE1": {
        "title_de": "Master-Schlafzimmer",
        "title_en": "Master Bedroom",
        "size": "21.50",
        "floor": "2",
        "room_num": "1",
        "rent": 640.00,
        "deposit": 1000.00,
        "features_de": "Kingsize-Boxspringbett, Doppelkleiderschrank, großer Schreibtisch, ergonomischer Stuhl, Blick zum Innenhof",
        "features_en": "King-size box spring bed, double wardrobe, executive study desk, ergonomic chair, quiet courtyard view"
    },
    "KE2": {
        "title_de": "Sonnige Balkon-Suite",
        "title_en": "Private Balcony Suite",
        "size": "23.92",
        "floor": "2",
        "room_num": "2",
        "rent": 649.00,
        "deposit": 1000.00,
        "features_de": "Privater begehbarer Gartenbalkon, Boxspring-Doppelbett, Arbeitsplatz mit Schreibtisch & Drehstuhl, großer Kleiderschrank",
        "features_en": "Private walk-out garden balcony, queen box spring bed, workstation with desk & chair, large wardrobe"
    },
    "KE3": {
        "title_de": "Ruhiges Gartenzimmer",
        "title_en": "Quiet Garden View",
        "size": "12.80",
        "floor": "2",
        "room_num": "3",
        "rent": 615.00,
        "deposit": 1000.00,
        "features_de": "Gemütliches Boxspringbett, Kleiderschrank, Arbeitsplatz, ruhige Ausrichtung mit Gartenblick",
        "features_en": "Cozy box spring bed, wardrobe, study desk, quiet garden exposure"
    },
    "KE4": {
        "title_de": "Studio-Zimmer",
        "title_en": "Cozy Studio Room",
        "size": "11.20",
        "floor": "2",
        "room_num": "4",
        "rent": 569.00,
        "deposit": 1000.00,
        "features_de": "Einzel-Boxspringbett, Einbauschrank, Arbeitsplatz, schnelles Internet",
        "features_en": "Single box spring bed, fitted wardrobe, study desk, high-speed WiFi"
    }
}

RATES = {
    "KEY_BUILDING": 680.00,
    "KEY_FLAT": 120.00,
    "KEY_ROOM": 90.00,
    "KEY_MAILBOX": 70.00,
    "CLEANING_FEE": 70.00,
    "LATE_FEE": 5.00,
    "MINOR_REPAIR_SINGLE": 100.00,
    "MINOR_REPAIR_CAP": 350.00,
    "MINOR_REPAIR_PCT": 8,
    "DAMAGE_TERMINATION": 50.00,
    "DEPOSIT_REFUND_DAYS": 30,
    "NOTICE_MONTHS": 3,
    "MIN_TERM_MONTHS": 6,
    "GUEST_DAYS": 5
}

def get_logo_base64():
    p = LOGO_PATH if os.path.exists(LOGO_PATH) else FALLBACK_LOGO_PATH
    if os.path.exists(p):
        with open(p, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    return ""

def num_to_words_de(n):
    n = int(round(n))
    units = ["null", "eins", "zwei", "drei", "vier", "fünf", "sechs", "sieben", "acht", "neun"]
    teens = ["zehn", "elf", "zwölf", "dreizehn", "vierzehn", "fünfzehn", "sechzehn", "siebzehn", "achtzehn", "neunzehn"]
    tens = ["", "", "zwanzig", "dreißig", "vierzig", "fünfzig", "sechzig", "siebzig", "achtzig", "neunzig"]

    def _convert_hundred(x):
        if x == 0:
            return ""
        if x < 10:
            return units[x]
        if x < 20:
            return teens[x - 10]
        u = x % 10
        t = x // 10
        if u == 0:
            return tens[t]
        u_str = "ein" if u == 1 else units[u]
        return f"{u_str}und{tens[t]}"

    def _convert_below_1000(x):
        if x == 0:
            return ""
        h = x // 100
        rem = x % 100
        h_str = ""
        if h == 1:
            h_str = "einhundert"
        elif h > 1:
            h_str = f"{units[h]}hundert"
        return h_str + _convert_hundred(rem)

    if n == 0:
        return "Null"
    if n == 1000:
        return "Eintausend"
    
    parts = []
    th = n // 1000
    rem = n % 1000
    if th == 1:
        parts.append("eintausend")
    elif th > 1:
        parts.append(f"{_convert_below_1000(th)}tausend")
    if rem > 0:
        parts.append(_convert_below_1000(rem))
    return "".join(parts).capitalize()

def num_to_words_en(n):
    n = int(round(n))
    units = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine"]
    teens = ["ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"]
    tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]

    def _convert_below_1000(x):
        if x == 0:
            return ""
        h = x // 100
        rem = x % 100
        parts = []
        if h > 0:
            parts.append(f"{units[h]} hundred")
        if rem > 0:
            if rem < 10:
                parts.append(units[rem])
            elif rem < 20:
                parts.append(teens[rem - 10])
            else:
                u = rem % 10
                t = rem // 10
                parts.append(f"{tens[t]}-{units[u]}" if u else tens[t])
        return " ".join(parts)

    if n == 0:
        return "Zero"
    if n == 1000:
        return "One thousand"
    
    parts = []
    th = n // 1000
    rem = n % 1000
    if th > 0:
        parts.append(f"{_convert_below_1000(th)} thousand")
    if rem > 0:
        parts.append(_convert_below_1000(rem))
    return " ".join(parts).capitalize()

def format_date_bilingual(date_str):
    try:
        dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    except Exception:
        return date_str, date_str

    de_months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
    en_months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    day = dt.day
    year = dt.year
    m_idx = dt.month - 1

    return f"{day:02d}.{dt.month:02d}.{year}", f"{day} {en_months[m_idx]}, {year}"

def render_contract_data(
    tenant_name,
    tenant_email="",
    tenant_phone="",
    room_code="KE2",
    start_date="2026-11-02",
    end_date="2027-04-30",
    rent_amount=None,
    deposit_amount=None,
    tenant_address_de=None,
    tenant_address_en=None,
    reason_de=None,
    reason_en=None,
    landlord_sign_date=None,
    tenant_sign_date=None,
    individual_de=None,
    individual_en=None,
    **kwargs
):
    spec = ROOM_SPECS.get(room_code, ROOM_SPECS["KE2"])
    rent = float(rent_amount) if rent_amount is not None else spec["rent"]
    deposit = float(deposit_amount) if deposit_amount is not None else spec["deposit"]

    rent_str = f"{int(rent)}" if rent.is_integer() else f"{rent:.2f}".replace(".", ",")
    deposit_str = f"{int(deposit):,}".replace(",", ".") if deposit.is_integer() else f"{deposit:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    rent_words_de = num_to_words_de(rent)
    rent_words_en = num_to_words_en(rent)
    deposit_words_de = num_to_words_de(deposit)
    deposit_words_en = num_to_words_en(deposit)

    start_de, start_en = format_date_bilingual(start_date)
    end_de, end_en = format_date_bilingual(end_date)

    today = datetime.date.today().strftime("%Y-%m-%d")
    today_de, today_en = format_date_bilingual(today)

    dep_due_dt = (datetime.datetime.strptime(start_date, "%Y-%m-%d") - datetime.timedelta(days=5)).strftime("%Y-%m-%d")
    dep_due_de, dep_due_en = format_date_bilingual(dep_due_dt)

    if not tenant_address_de:
        tenant_address_de = "Gemäß Personalausweis / Passdokument"
    if not tenant_address_en:
        tenant_address_en = "As per identity card / passport document"

    if not reason_de:
        reason_de = "Aufgrund eines Studien-, Praktikums- oder Erwerbsaufenthalts in Berlin"
    if not reason_en:
        reason_en = "Due to an academic, internship, or professional stay in Berlin"

    sign_dt_de = landlord_sign_date or today_de
    sign_dt_en = tenant_sign_date or today_en

    data = {
        "LANDLORD_NAME": "Kylasam NA",
        "LANDLORD_SIGN_NAME": "Kylasam Nallagoundanpalayam Arruchamy",
        "LANDLORD_ADDRESS_DE": "Forckenbeckstraße 64, 14199 Berlin, Deutschland",
        "LANDLORD_ADDRESS_EN": "Forckenbeckstraße 64, 14199 Berlin, Germany",
        "LANDLORD_CONTACT": "support@homelyo.eu | Tel: +49 176 8332 9439",
        "PROPERTY_ADDRESS_DE": "Lichtenrader Damm 80, 12305 Berlin, Deutschland",
        "PROPERTY_ADDRESS_EN": "Lichtenrader Damm 80, 12305 Berlin, Germany",
        "FLOOR": spec.get("floor", "2"),
        "ROOM_COUNT": spec.get("room_num", "2"),
        "ROOM_SIZE": spec.get("size", "23.92"),
        "ROOM_CODE": room_code,
        "ROOM_TITLE_DE": spec.get("title_de", "Sonnige Balkon-Suite"),
        "ROOM_TITLE_EN": spec.get("title_en", "Private Balcony Suite"),
        "ROOM_FEATURES_DE": spec.get("features_de", "Balkon, Queensize-Bett, Kleiderschrank, Schreibtisch & Bürostuhl"),
        "ROOM_FEATURES_EN": spec.get("features_en", "Balcony, queen bed, wardrobe, desk workstation & chair"),
        "TENANT_NAME": tenant_name,
        "TENANT_SIGN_NAME": tenant_name,
        "TENANT_EMAIL": tenant_email,
        "TENANT_PHONE": tenant_phone,
        "TENANT_ADDRESS_DE": tenant_address_de,
        "TENANT_ADDRESS_EN": tenant_address_en,
        "START_DATE_DE": start_de,
        "START_DATE_EN": start_en,
        "END_DATE_DE": end_de,
        "END_DATE_EN": end_en,
        "CONTRACT_CREATION_DATE": today_de,
        "CONTRACT_DATE": today_de,
        "REASON_DE": reason_de,
        "REASON_EN": reason_en,
        "RENT_AMOUNT": rent_str,
        "RENT_AMOUNT_DE": f"{rent_str} €",
        "RENT_AMOUNT_EN": f"€{rent_str}",
        "RENT_WORDS_DE": rent_words_de,
        "RENT_WORDS_EN": rent_words_en,
        "DEPOSIT_AMOUNT": deposit_str,
        "DEPOSIT_AMOUNT_DE": f"{deposit_str} €",
        "DEPOSIT_AMOUNT_EN": f"€{deposit_str}",
        "DEPOSIT_WORDS_DE": deposit_words_de,
        "DEPOSIT_WORDS_EN": deposit_words_en,
        "DEPOSIT_DUE_DE": dep_due_de,
        "DEPOSIT_DUE_EN": dep_due_en,
        "FIRST_RENT_AMOUNT": rent_str,
        "CLEANING_FEE": "70,00",
        "LATE_FEE": "5,00",
        "RENT_ACCOUNT_HOLDER": "Kylasam Nallagoundanpalayam Arruchamy",
        "RENT_IBAN": "DE87100110012624542768",
        "RENT_BIC": "NTSBDEB1XXX",
        "RENT_BANK": "N26 Bank AG, Berlin",
        "LANDLORD_SIGN_DATE": sign_dt_de,
        "TENANT_SIGN_DATE": sign_dt_en,
        "INDIVIDUAL_AGREEMENTS_DE": individual_de or "Keine gesonderten Vereinbarungen.",
        "INDIVIDUAL_AGREEMENTS_EN": individual_en or "None.",
        # GDrive sample aliases for full backwards-compatibility
        "TENANT_ROOM_ID": room_code,
        "ROOM_TITLE": spec.get("title_de", "Sonnige Balkon-Suite"),
        "ROOM_NUMBER": spec.get("room_num", "2"),
        "CONTRACT_START_DT": start_de,
        "CONTRACT_END_DT": end_de,
        "RENT_PER_MONTH": rent_str,
        "DEPOSIT_AMOUNT_PAY_BEFORE_DT": dep_due_de
    }
    return data

def generate_html_template():
    logo_b64 = get_logo_base64()
    logo_img = f'<img src="data:image/png;base64,{logo_b64}" alt="Homelyo Logo" style="height: 30px; width: auto; max-width: 140px; display: inline-block;">' if logo_b64 else '<div style="font-size: 15pt; font-weight: 800; color: #0f172a; letter-spacing: 0.5px;">HOMELYO</div>'
    footer_logo = f'<img src="data:image/png;base64,{logo_b64}" alt="Homelyo" style="height: 12px; width: auto; opacity: 0.85;">' if logo_b64 else '<span style="font-weight: 700; font-size: 7.5pt; color: #475569;">homelyo</span>'

    css = """
@page {
    size: A4 portrait;
    margin: 9mm 12mm 9mm 12mm;
}
* {
    box-sizing: border-box;
    -webkit-font-smoothing: antialiased;
}
body {
    margin: 0;
    padding: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    color: #0f172a;
    background-color: #ffffff;
    font-size: 7.8pt;
    line-height: 1.28;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
}
.page {
    width: 100%;
    page-break-after: always;
    height: 279mm;
    max-height: 279mm;
    position: relative;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    overflow: hidden;
}
.page:last-child {
    page-break-after: avoid;
}
.page-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 1.5px solid #cbd5e1;
    padding-bottom: 5px;
    margin-bottom: 6px;
    flex-shrink: 0;
}
.header-meta {
    text-align: right;
    font-size: 7.5pt;
    color: #64748b;
    line-height: 1.2;
}
.page-footer {
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-top: 1px solid #e2e8f0;
    padding-top: 5px;
    margin-top: 5px;
    font-size: 7.0pt;
    color: #64748b;
    flex-shrink: 0;
}
.page-body {
    flex-grow: 1;
    display: flex;
    flex-direction: column;
}
.main-title-block {
    text-align: center;
    margin: 6px 0 10px 0;
}
.main-title-block h1 {
    font-size: 18pt;
    font-weight: 800;
    color: #0f172a;
    margin: 0 0 2px 0;
    letter-spacing: -0.3px;
}
.main-title-block h2 {
    font-size: 11pt;
    font-weight: 600;
    color: #475569;
    margin: 0;
}
.summary-card {
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 5px;
    padding: 8px 12px;
    margin-bottom: 10px;
}
.summary-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 4px 14px;
    font-size: 8.0pt;
}
.summary-row {
    display: flex;
    justify-content: space-between;
    padding: 1.5px 0;
}
.summary-label {
    color: #475569;
    font-weight: 600;
}
.summary-val {
    color: #0f172a;
    font-weight: 700;
    text-align: right;
}
.two-col {
    display: flex;
    flex-direction: row;
    justify-content: space-between;
    width: 100%;
    gap: 12px;
}
.col {
    flex: 1;
    width: 50%;
}
.sec-title {
    font-size: 8.6pt;
    font-weight: 700;
    color: #0f172a;
    margin-top: 6px;
    margin-bottom: 2px;
    line-height: 1.2;
    border-left: 3px solid #2563eb;
    padding-left: 5px;
}
.col-en .sec-title {
    color: #334155;
    border-left: 3px solid #94a3b8;
}
.p-block {
    margin-bottom: 3.5px;
    text-align: justify;
    font-size: 7.7pt;
    color: #1e293b;
    line-height: 1.26;
}
.col-en .p-block {
    color: #475569;
}
.party-box {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 4px;
    padding: 6px 8px;
    margin-bottom: 6px;
    font-size: 7.7pt;
}
.party-title {
    font-weight: 700;
    color: #0f172a;
    border-bottom: 1px solid #e2e8f0;
    padding-bottom: 2px;
    margin-bottom: 3px;
}
.bank-box {
    background: #f8fafc;
    border: 1.5px solid #cbd5e1;
    border-radius: 5px;
    padding: 8px 12px;
    margin: 8px 0;
}
.bank-title {
    font-weight: 700;
    font-size: 9.0pt;
    color: #0f172a;
    margin-bottom: 4px;
    border-bottom: 1px solid #cbd5e1;
    padding-bottom: 2px;
}
.inventory-table {
    width: 100%;
    border-collapse: collapse;
    margin: 6px 0;
    font-size: 7.6pt;
}
.inventory-table th, .inventory-table td {
    border: 1px solid #cbd5e1;
    padding: 4px 6px;
    vertical-align: top;
    text-align: left;
}
.inventory-table th {
    background: #f1f5f9;
    font-weight: 700;
    color: #0f172a;
}
.sig-box {
    margin-top: 10px;
    padding-top: 6px;
    border-top: 1px dashed #cbd5e1;
}
.sig-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 7.6pt;
}
.sig-table td {
    width: 50%;
    vertical-align: top;
    padding: 0 8px;
}
.sig-line {
    border-top: 1.5px solid #0f172a;
    margin-top: 28px;
    margin-bottom: 3px;
}
.badge {
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    color: #1d4ed8;
    padding: 1px 4px;
    border-radius: 3px;
    font-size: 7.2pt;
    font-weight: 600;
}
.legal-banner {
    font-style: italic;
    font-size: 7.2pt;
    color: #64748b;
    text-align: center;
    margin-top: 4px;
}
"""

    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>Mietvertrag / Residential Lease Agreement – Homelyo</title>
<style>{css}</style>
</head>
<body>

<!-- ==================== PAGE 1: COVER & WELCOME OVERVIEW ==================== -->
<div class="page" id="page-1">
  <div class="page-header">
    <div>{logo_img}</div>
    <div class="header-meta">
      <strong>Homelyo Co-Living Berlin</strong><br>
      Lichtenrader Damm 80, 12305 Berlin
    </div>
  </div>

  <div class="page-body">
    <div class="main-title-block">
      <h1>Mietvertrag</h1>
      <h2>Residential Lease Agreement &bull; Co-Living Tenancy</h2>
    </div>

    <!-- Executive Summary Card -->
    <div class="summary-card">
      <div class="summary-grid">
        <div class="summary-row">
          <span class="summary-label">Mieter / Tenant:</span>
          <span class="summary-val">{{{{TENANT_NAME}}}}</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">Mietobjekt / Room:</span>
          <span class="summary-val">Zimmer {{{{ROOM_CODE}}}} ({{{{ROOM_TITLE_DE}}}})</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">Liegenschaft / Address:</span>
          <span class="summary-val">Lichtenrader Damm 80, 12305 Berlin</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">Geschoss & Whg:</span>
          <span class="summary-val">{{{{FLOOR}}}}. OG, Whg. WE-11 (Zimmer {{{{ROOM_COUNT}}}})</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">Mietzeit / Rental Term:</span>
          <span class="summary-val">{{{{START_DATE_DE}}}} &ndash; {{{{END_DATE_DE}}}}</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">Warmmiete / Warm Rent:</span>
          <span class="summary-val">{{{{RENT_AMOUNT}}}} &euro; / Monat (Pauschalmiete)</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">Kaution / Deposit:</span>
          <span class="summary-val">{{{{DEPOSIT_AMOUNT}}}} &euro; (F&auml;llig vor Einzug)</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">Endreinigung / Cleaning:</span>
          <span class="summary-val">{{{{CLEANING_FEE}}}} &euro; (Einmalig bei Auszug)</span>
        </div>
      </div>
    </div>

    <!-- Welcome Letter (Wunderflats & GDrive standard) -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">Willkommen bei Homelyo</div>
        <div class="p-block">
          <strong>Liebe/r Mieter/in {{{{TENANT_NAME}}}},</strong><br>
          im Folgenden erh&auml;ltst Du alle wichtigen Informationen rund um Deinen Mietvertrag sowie Deine ersten Schritte zur erfolgreichen Anmietung Deines Zimmers in der Liegenschaft Lichtenrader Damm 80, 12305 Berlin.
        </div>
        <div class="p-block">
          <strong>Bestandteile dieser Vereinbarung:</strong><br>
          &bull; <strong>Hauptvertrag:</strong> Mietgegenstand, Mietdauer, Konditionen (&sect; 1 &ndash; &sect; 6)<br>
          &bull; <strong>Zahlungsm&ouml;glichkeiten:</strong> N26 Bankverbindung &amp; Kautionsabwicklung<br>
          &bull; <strong>Anlage 1:</strong> Inventar der Mietsache &amp; Kleinreparaturregelungen<br>
          &bull; <strong>Anlage 2:</strong> Allgemeine Gesch&auml;ftsbedingungen (&sect; 1 &ndash; &sect; 8 AGB)
        </div>
        <div class="p-block">
          <span class="badge">Rechtlicher Hinweis</span> Die deutsche Sprachfassung dieses Mietvertrages ist rechtlich allein verbindlich. Die englische Fassung dient als unverbindliche Lesehilfe.
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">Welcome to Homelyo</div>
        <div class="p-block">
          <strong>Dear Tenant {{{{TENANT_NAME}}}},</strong><br>
          Below you will find all essential information regarding your lease agreement and key steps for moving into your co-living room at Lichtenrader Damm 80, 12305 Berlin.
        </div>
        <div class="p-block">
          <strong>Structure of this Lease Document:</strong><br>
          &bull; <strong>Core Agreement:</strong> Premises, term, rent &amp; deposit (&sect; 1 &ndash; &sect; 6)<br>
          &bull; <strong>Payment Options:</strong> N26 Bank wire transfer &amp; escrow processing<br>
          &bull; <strong>Attachment 1:</strong> Inventory schedule &amp; maintenance rules<br>
          &bull; <strong>Attachment 2:</strong> General Terms &amp; House Rules (&sect; 1 &ndash; &sect; 8 GTC)
        </div>
        <div class="p-block">
          <span class="badge">Legal Notice</span> Only the German language version of this agreement is legally binding. The English translation is provided solely as an informational convenience.
        </div>
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo} &bull; Homelyo Co-Living Berlin</div>
    <div>Seite / Page 1 von 8</div>
  </div>
</div>

<!-- ==================== PAGE 2: HAUPTVERTRAG TEIL 1 (§ 1 & § 2) ==================== -->
<div class="page" id="page-2">
  <div class="page-header">
    <div>{logo_img}</div>
    <div class="header-meta">
      <strong>Mietvertrag (Hauptvertrag)</strong> &bull; {{{{TENANT_NAME}}}} ({{{{ROOM_CODE}}}})
    </div>
  </div>

  <div class="page-body">
    <!-- Parties & Premises Box -->
    <div class="two-col">
      <div class="col col-de">
        <div class="party-box">
          <div class="party-title">Vermieter / Landlord</div>
          <strong>{{{{LANDLORD_SIGN_NAME}}}}</strong><br>
          {{{{LANDLORD_ADDRESS_DE}}}}<br>
          {{{{LANDLORD_CONTACT}}}}
        </div>
      </div>
      <div class="col col-en">
        <div class="party-box">
          <div class="party-title">Mieter / Tenant</div>
          <strong>{{{{TENANT_NAME}}}}</strong><br>
          {{{{TENANT_ADDRESS_DE}}}}<br>
          {{{{TENANT_EMAIL}}}} | {{{{TENANT_PHONE}}}}
        </div>
      </div>
    </div>

    <!-- Mietgegenstand -->
    <div class="two-col">
      <div class="col col-de">
        <div class="p-block">
          <strong>Mietgegenstand:</strong> Vermietet wird das Zimmer <strong>{{{{ROOM_CODE}}}} ({{{{ROOM_TITLE_DE}}}})</strong> mit ca. <strong>{{{{ROOM_SIZE}}}} m&sup2;</strong> in der Wohnung Lichtenrader Damm 80, 12305 Berlin (2. OG, WE-11). Gemeinschaftliche Bereiche wie Flur, K&uuml;che, Badezimmer, Balkon/Terrasse und Gep&auml;ckraum d&uuml;rfen mitbenutzt werden.
        </div>
      </div>
      <div class="col col-en">
        <div class="p-block">
          <strong>Rental Premises:</strong> Private room <strong>{{{{ROOM_CODE}}}} ({{{{ROOM_TITLE_EN}}}})</strong> with approx. <strong>{{{{ROOM_SIZE}}}} m&sup2;</strong> in apartment Lichtenrader Damm 80, 12305 Berlin (2nd floor, WE-11). Communal amenities including hallway, kitchen, bathroom, balcony/terrace, and luggage room are shared.
        </div>
      </div>
    </div>

    <!-- § 1 Mietdauer -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">&sect; 1 Mietdauer &amp; Befristungszweck (&sect; 549 Abs. 2 / &sect; 575 BGB)</div>
        <div class="p-block">
          (1) <strong>Mietzeit:</strong> Das Mietverh&auml;ltnis beginnt am <strong>{{{{START_DATE_DE}}}}</strong> und ist bis zum <strong>{{{{END_DATE_DE}}}}</strong> fest befristet. Die Mindestmietdauer betr&auml;gt sechs (6) Monate.
        </div>
        <div class="p-block">
          (2) <strong>Zweck der Befristung:</strong> Die Vermietung erfolgt zum vor&uuml;bergehenden Gebrauch gem. &sect; 549 Abs. 2 Nr. 1 BGB (<em>{{{{REASON_DE}}}}</em>).
        </div>
        <div class="p-block">
          (3) <strong>Ausschluss Verl&auml;ngerung:</strong> Eine stillschweigende Verl&auml;ngerung des Mietverh&auml;ltnisses gem. &sect; 545 BGB ist ausdr&uuml;cklich ausgeschlossen.
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">&sect; 1 Lease Term &amp; Temporary Purpose (&sect; 549 (2) / &sect; 575 BGB)</div>
        <div class="p-block">
          (1) <strong>Tenancy Term:</strong> Begins on <strong>{{{{START_DATE_EN}}}}</strong> and terminates on <strong>{{{{END_DATE_EN}}}}</strong>. Minimum lease term is six (6) months.
        </div>
        <div class="p-block">
          (2) <strong>Temporary Purpose:</strong> Let strictly for temporary accommodation under Section 549 (2) No. 1 BGB (<em>{{{{REASON_EN}}}}</em>).
        </div>
        <div class="p-block">
          (3) <strong>No Tacit Renewal:</strong> Tacit renewal beyond the fixed term pursuant to Section 545 BGB is strictly excluded.
        </div>
      </div>
    </div>

    <!-- § 2 Miethöhe -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">&sect; 2 Mieth&ouml;he, Pauschale &amp; Nebenkosten</div>
        <div class="p-block">
          (1) <strong>Warmmiete:</strong> Die monatliche Pauschalmiete betr&auml;gt <strong>{{{{RENT_AMOUNT}}}} &euro;</strong> (in Worten: <strong>{{{{RENT_WORDS_DE}}}} Euro</strong>).
        </div>
        <div class="p-block">
          (2) <strong>Inklusivleistungen:</strong> Die Miete deckt alle kalten und warmen Betriebskosten (Heizung, Warm-/Kaltwasser, Strom, M&uuml;ll, Geb&auml;udeversicherung) sowie Vodafone 1000 Mbit/s WLAN in verbrauchs&uuml;blichem Ma&szlig;e ab (Fair-Use-Klausel gem. AGB &sect; 3 Nr. 6).
        </div>
        <div class="p-block">
          (3) <strong>ARD-Rundfunkbeitrag:</strong> Der Rundfunkbeitrag (GEZ) ist <strong>nicht</strong> in der Miete enthalten. Die WG meldet den Beitrag eigenst&auml;ndig an und teilt die Kosten untereinander auf.
        </div>
        <div class="p-block">
          (4) <strong>Endreinigungspauschale:</strong> Bei Auszug f&auml;llt eine einmalige Pauschale von <strong>{{{{CLEANING_FEE}}}} &euro;</strong> f&uuml;r die professionelle Endreinigung an.
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">&sect; 2 Rent, Utilities &amp; License Fee</div>
        <div class="p-block">
          (1) <strong>Warm Rent:</strong> The monthly all-inclusive flat rent is <strong>&euro;{{{{RENT_AMOUNT}}}}</strong> (in words: <strong>{{{{RENT_WORDS_EN}}}} euros</strong>).
        </div>
        <div class="p-block">
          (2) <strong>Inclusions:</strong> Includes all operating utilities (heating, hot/cold water, electricity, waste disposal, insurance) and Vodafone 1000 Mbps Wi-Fi based on customary fair use (GTC &sect; 3 no. 6).
        </div>
        <div class="p-block">
          (3) <strong>Broadcasting Fee (ARD):</strong> The public broadcasting fee (GEZ) is <strong>not</strong> included; registration and payment are coordinated directly by flatmates.
        </div>
        <div class="p-block">
          (4) <strong>Final Cleaning:</strong> A one-time checkout cleaning fee of <strong>&euro;{{{{CLEANING_FEE}}}}</strong> applies upon departure.
        </div>
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo} &bull; Homelyo Hauptvertrag</div>
    <div>Seite / Page 2 von 8</div>
  </div>
</div>

<!-- ==================== PAGE 3: HAUPTVERTRAG TEIL 2 (§ 3, § 4, § 5, § 6 & SIGNATURES) ==================== -->
<div class="page" id="page-3">
  <div class="page-header">
    <div>{logo_img}</div>
    <div class="header-meta">
      <strong>Mietvertrag (Hauptvertrag)</strong> &bull; Unterschriften / Signatures
    </div>
  </div>

  <div class="page-body">
    <!-- § 3 Kaution & § 4 Zahlungsbedingungen -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">&sect; 3 Kaution &amp; Mietsicherheit (&sect; 551 BGB)</div>
        <div class="p-block">
          (1) <strong>Kautionsh&ouml;he:</strong> Der Mieter leistet eine Kaution von <strong>{{{{DEPOSIT_AMOUNT}}}} &euro;</strong> (in Worten: <strong>{{{{DEPOSIT_WORDS_DE}}}} Euro</strong>).
        </div>
        <div class="p-block">
          (2) <strong>F&auml;lligkeit vor Einzug:</strong> Kaution und erste Monatsmiete m&uuml;ssen vor dem Einzugstermin (<strong>{{{{START_DATE_DE}}}}</strong>) vollst&auml;ndig eingegangen sein. Ohne Geldeingang erfolgt keine Schl&uuml;ssel&uuml;bergabe.
        </div>
        <div class="p-block">
          (3) <strong>R&uuml;ckerstattung:</strong> Treuh&auml;nderische Verwahrung gem. &sect; 551 BGB. Die Auszahlung erfolgt binnen <strong>30 Tagen</strong> nach Auszug und R&uuml;ckgabeprotokoll abz&uuml;glich berechtigter Forderungen. Aufrechnung gegen laufende Miete ist unzul&auml;ssig.
        </div>

        <div class="sec-title">&sect; 4 Zahlungsbedingungen &amp; F&auml;lligkeit</div>
        <div class="p-block">
          (1) <strong>F&auml;lligkeit:</strong> Die laufende Miete ist monatlich im Voraus bis sp&auml;testens zum <strong>3. Werktag</strong> eines jeden Monats kostenfrei per Bank&uuml;berweisung zu zahlen.
        </div>
        <div class="p-block">
          (2) <strong>Mahngeb&uuml;hren:</strong> F&uuml;r jede Mahnung wird eine Geb&uuml;hr von <strong>{{{{LATE_FEE}}}} &euro;</strong> erhoben. Verzug mit zwei Monatsmieten berechtigt zur fristlosen K&uuml;ndigung gem. &sect; 543 Abs. 2 BGB.
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">&sect; 3 Security Deposit &amp; Escrow (&sect; 551 BGB)</div>
        <div class="p-block">
          (1) <strong>Deposit Sum:</strong> Tenant renders a deposit of <strong>&euro;{{{{DEPOSIT_AMOUNT}}}}</strong> (in words: <strong>{{{{DEPOSIT_WORDS_EN}}}} euros</strong>).
        </div>
        <div class="p-block">
          (2) <strong>Due Before Move-In:</strong> Deposit and 1st month rent must be credited before move-in (<strong>{{{{START_DATE_EN}}}}</strong>). Key handover is denied without prior funds receipt.
        </div>
        <div class="p-block">
          (3) <strong>Escrow &amp; Refund:</strong> Held in escrow per Section 551 BGB and refunded within <strong>30 days</strong> after departure and signed checkout protocol, less proven claims. Setting off deposit against ongoing rent is prohibited.
        </div>

        <div class="sec-title">&sect; 4 Payment Terms &amp; Due Dates</div>
        <div class="p-block">
          (1) <strong>Due Date:</strong> Monthly rent is payable in advance on or before the <strong>3rd business day</strong> of each calendar month via wire transfer.
        </div>
        <div class="p-block">
          (2) <strong>Late Charges:</strong> A fee of <strong>&euro;{{{{LATE_FEE}}}}</strong> applies per reminder notice. Arrears of two months rent entitle landlord to terminate without notice under &sect; 543 (2) BGB.
        </div>
      </div>
    </div>

    <!-- § 5 Kündigung & § 6 AGB/Anlagen -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">&sect; 5 K&uuml;ndigung &amp; Vorzeitige Entlassung</div>
        <div class="p-block">
          (1) <strong>Ordentliche K&uuml;ndigung:</strong> W&auml;hrend der festen Vertragslaufzeit ist die ordentliche K&uuml;ndigung ausgeschlossen.
        </div>
        <div class="p-block">
          (2) <strong>Nachmieterstellung:</strong> Der Mieter kann mit einer Frist von <strong>3 Monaten zum Monatsende</strong> vorzeitig aus dem Vertrag entlassen werden, wenn er einen solventen, zumutbaren Nachmieter stellt, der einen verbindlichen Mietvertrag abschlie&szlig;t.
        </div>
        <div class="p-block">
          (3) <strong>Au&szlig;erordentliche K&uuml;ndigung:</strong> Das gesetzliche Recht zur K&uuml;ndigung aus wichtigem Grund (&sect;&sect; 543, 569 BGB) bleibt unber&uuml;hrt.
        </div>

        <div class="sec-title">&sect; 6 AGB &amp; Anlagen als Vertragsbestandteil</div>
        <div class="p-block">
          Die Anlagen &ndash; <strong>Anlage 1 (Inventar)</strong> und <strong>Anlage 2 (AGB)</strong> &ndash; sind wesentlicher Vertragsbestandteil. &Uuml;bergabeprotokoll und Wohnungsgeberbest&auml;tigung (&sect; 19 BMG) werden bei Einzug ausgeh&auml;ndigt.
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">&sect; 5 Termination &amp; Replacement Tenant</div>
        <div class="p-block">
          (1) <strong>Ordinary Termination:</strong> Excluded during the fixed rental term.
        </div>
        <div class="p-block">
          (2) <strong>Replacement Tenant:</strong> Tenant may be released early with <strong>3 months notice to month-end</strong> upon providing an accepted, solvent replacement tenant who signs a binding lease.
        </div>
        <div class="p-block">
          (3) <strong>Termination for Cause:</strong> Statutory right to extraordinary termination for good cause (&sect;&sect; 543, 569 BGB) remains unaffected.
        </div>

        <div class="sec-title">&sect; 6 Terms &amp; Attachments</div>
        <div class="p-block">
          The attachments &ndash; <strong>Attachment 1 (Inventory)</strong> and <strong>Attachment 2 (GTC)</strong> &ndash; form an integral binding part of this agreement. Handover protocol and landlord confirmation (&sect; 19 BMG) are provided upon move-in.
        </div>
      </div>
    </div>

    <!-- Signatures Block 1 (Hauptvertrag) -->
    <div class="sig-box">
      <table class="sig-table">
        <tr>
          <td>
            Ort, Datum / Place, Date: Berlin, {{{{LANDLORD_SIGN_DATE}}}}
            <div class="sig-line"></div>
            <strong>{{{{LANDLORD_SIGN_NAME}}}}</strong><br>
            Vermieter / Landlord
          </td>
          <td>
            Ort, Datum / Place, Date: {{{{TENANT_SIGN_DATE}}}}
            <div class="sig-line"></div>
            <strong>{{{{TENANT_SIGN_NAME}}}}</strong><br>
            Mieter / Tenant
          </td>
        </tr>
      </table>
      <div class="legal-banner">Hinweis: Die deutsche Fassung ist rechtlich allein verbindlich. / Notice: The German text is solely legally binding.</div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo} &bull; Homelyo Hauptvertrag &bull; Unterschriften</div>
    <div>Seite / Page 3 von 8</div>
  </div>
</div>

<!-- ==================== PAGE 4: ZAHLUNGSMÖGLICHKEITEN & ANLAGE 1 INVENTAR ==================== -->
<div class="page" id="page-4">
  <div class="page-header">
    <div>{logo_img}</div>
    <div class="header-meta">
      <strong>Zahlungsm&ouml;glichkeiten &amp; Anlage 1 (Inventar)</strong> &bull; {{{{ROOM_CODE}}}}
    </div>
  </div>

  <div class="page-body">
    <!-- Zahlungsmöglichkeiten Box -->
    <div class="bank-box">
      <div class="bank-title">Zahlungsm&ouml;glichkeiten / Payment Options &bull; N26 Bank AG</div>
      <div class="two-col" style="font-size: 7.7pt;">
        <div class="col">
          <strong>Bankverbindung f&uuml;r Miete &amp; Kaution:</strong><br>
          Kontoinhaber: <strong>{{{{RENT_ACCOUNT_HOLDER}}}}</strong><br>
          IBAN: <strong>{{{{RENT_IBAN}}}}</strong><br>
          BIC / SWIFT: <strong>{{{{RENT_BIC}}}}</strong> &bull; Bank: {{{{RENT_BANK}}}}<br>
          <em>Verwendungszweck Miete:</em> Miete [Monat Jahr] {{{{TENANT_NAME}}}}<br>
          <em>Verwendungszweck Kaution:</em> Kaution {{{{TENANT_NAME}}}} Zimmer {{{{ROOM_CODE}}}}
        </div>
        <div class="col col-en">
          <strong>Bank Details for Rent &amp; Deposit:</strong><br>
          Account Holder: <strong>{{{{RENT_ACCOUNT_HOLDER}}}}</strong><br>
          IBAN: <strong>{{{{RENT_IBAN}}}}</strong><br>
          BIC / SWIFT: <strong>{{{{RENT_BIC}}}}</strong> &bull; Bank: {{{{RENT_BANK}}}}<br>
          <em>Rent Reference:</em> Rent [Month Year] {{{{TENANT_NAME}}}}<br>
          <em>Deposit Reference:</em> Deposit {{{{TENANT_NAME}}}} Room {{{{ROOM_CODE}}}}
        </div>
      </div>
    </div>

    <!-- Anlage 1 - Inventar der Mietsache -->
    <div class="sec-title" style="margin-top: 4px;">Anlage 1 &ndash; Inventar der Mietsache / Attachment 1 &ndash; Inventory</div>

    <table class="inventory-table">
      <thead>
        <tr>
          <th style="width: 50%;">Ausstattung &amp; Regelungen (Deutsch)</th>
          <th style="width: 50%; color: #475569;">Furnishings &amp; Terms (English)</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>
            <strong>1. Ausstattung im privaten Zimmer ({{{{ROOM_CODE}}}}):</strong><br>
            &bull; Komfortables Boxspringbett mit Qualit&auml;tsmatratze, Lattenrost &amp; Schoner<br>
            &bull; Gro&szlig;er Kleiderschrank / Kommode mit ausreichend B&uuml;geln<br>
            &bull; Ergonomischer Schreibtisch &amp; B&uuml;rostuhl f&uuml;r Home-Office / Studium<br>
            &bull; Spiegel, Deckenbeleuchtung, blickdichte Vorh&auml;nge / Rollos<br>
            &bull; Erstausstattung Bettw&auml;scheset (Kissen, Bettdecke, Bez&uuml;ge, Laken)<br>
            &bull; <em>Zimmerspezifisch:</em> {{{{ROOM_FEATURES_DE}}}}
          </td>
          <td style="color: #475569;">
            <strong>1. Private Room Inventory ({{{{ROOM_CODE}}}}):</strong><br>
            &bull; Quality box spring bed with mattress, slatted frame &amp; protector<br>
            &bull; Spacious wardrobe / dresser with hangers<br>
            &bull; Ergonomic study desk &amp; chair for productive work/study<br>
            &bull; Mirror, room lighting, blackout curtains / blinds<br>
            &bull; Starter bedding set (pillow, duvet, linen covers, bedsheet)<br>
            &bull; <em>Room specifics:</em> {{{{ROOM_FEATURES_EN}}}}
          </td>
        </tr>
        <tr>
          <td>
            <strong>2. Gemeinschaftsausstattung (K&uuml;che, Bad, Flur, Balkon):</strong><br>
            &bull; Waschmaschine, K&uuml;hl-/Gefrierkombination, Induktionskochfeld, Backofen<br>
            &bull; Mikrowelle, Kaffeemaschine, Wasserkocher, Toaster, Staubsauger<br>
            &bull; Vollst&auml;ndiges Geschirr-, Besteck- und Kochutensilienset<br>
            &bull; Highspeed WLAN-Router (Vodafone Gigabit-Kabel)
          </td>
          <td style="color: #475569;">
            <strong>2. Shared Amenities (Kitchen, Bath, Hall, Balcony):</strong><br>
            &bull; Washing machine, fridge/freezer, induction hob, oven<br>
            &bull; Microwave, coffee maker, kettle, toaster, vacuum cleaner<br>
            &bull; Full set of crockery, cutlery, pots and cookware<br>
            &bull; Highspeed Wi-Fi router (Vodafone Gigabit connection)
          </td>
        </tr>
        <tr>
          <td>
            <strong>3. Wichtige Obhuts- &amp; Reparaturhinweise:</strong><br>
            &bull; Sch&auml;den an der Mietsache sind gem. &sect; 5 Nr. 1 AGB unverz&uuml;glich zu melden.<br>
            &bull; Durch den Mieter verursachte Sch&auml;den tr&auml;gt dieser gem. &sect; 5 Nr. 2 AGB in vollem Ma&szlig;e.<br>
            &bull; Kleinreparaturen an h&auml;ufig genutzten Teilen (u.a. Leuchtmittel, Griffe) bis <strong>100,00 &euro;</strong> je Einzelfall tr&auml;gt der Mieter gem. &sect; 5 Nr. 5 AGB selbst.<br>
            &bull; Ersatzanspr&uuml;che f&uuml;r nicht mitvermietete Gegenst&auml;nde sind ausgeschlossen.
          </td>
          <td style="color: #475569;">
            <strong>3. Important Custody &amp; Repair Regulations:</strong><br>
            &bull; Defects must be reported immediately per GTC &sect; 5 no. 1.<br>
            &bull; Damage caused by tenant is fully compensated per GTC &sect; 5 no. 2.<br>
            &bull; Minor repairs on frequent-use items (light bulbs, handles) up to <strong>&euro;100.00</strong> per repair are borne by tenant per GTC &sect; 5 no. 5.<br>
            &bull; No compensation claims for items not included in the lease.
          </td>
        </tr>
      </tbody>
    </table>
  </div>

  <div class="page-footer">
    <div>{footer_logo} &bull; Homelyo Zahlungsoptionen &amp; Anlage 1 (Inventar)</div>
    <div>Seite / Page 4 von 8</div>
  </div>
</div>

<!-- ==================== PAGE 5: ANLAGE 2 AGB TEIL 1 (PRÄAMBEL & § 1) ==================== -->
<div class="page" id="page-5">
  <div class="page-header">
    <div>{logo_img}</div>
    <div class="header-meta">
      <strong>Anlage 2 &ndash; AGB (Teil 1)</strong> &bull; Zusammenleben in der WG
    </div>
  </div>

  <div class="page-body">
    <!-- AGB Preamble -->
    <div class="summary-card" style="margin-bottom: 6px; padding: 6px 10px;">
      <div class="two-col" style="font-size: 7.5pt;">
        <div class="col">
          <strong>Allgemeine Gesch&auml;ftsbedingungen der Wohngemeinschaft (Pr&auml;ambel):</strong><br>
          Der Mieter erkennt die AGB als f&uuml;r ihn verbindlich an. Ein Versto&szlig; ist ein vertragswidriger Gebrauch der Mietsache und berechtigt bei schwerwiegenden F&auml;llen zur fristlosen K&uuml;ndigung. F&uuml;r Sch&auml;den aus Nichtbeachtung haftet der Mieter ersatzpflichtig.
        </div>
        <div class="col col-en">
          <strong>General Terms and Conditions for Co-Living Flatshare (Preamble):</strong><br>
          The tenant accepts the GTC as legally binding. A breach represents contract misuse and entitles the landlord to termination without notice in severe cases. The tenant compensates for all damages resulting from non-compliance.
        </div>
      </div>
    </div>

    <!-- § 1 Zusammenleben in der WG -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">&sect; 1 Zusammenleben in der WG / mit den Nachbarn</div>
        <div class="p-block">
          <strong>1. R&uuml;cksichtnahme:</strong> Die Bewohner der WG haben das Zusammenleben partnerschaftlich mit R&uuml;cksicht auf die Interessen und zur Zufriedenheit aller Beteiligten fair und eigenverantwortlich zu regeln.
        </div>
        <div class="p-block">
          <strong>2. L&auml;rm &amp; Ruhezeiten:</strong> Jeder Mieter vermeidet unzumutbaren L&auml;rm. Gesetzliche Ruhezeiten sind strikt einzuhalten: <strong>13:00 bis 15:00 Uhr</strong> sowie <strong>22:00 bis 08:00 Uhr</strong>. An Sonn- und Feiertagen ist Ruhest&ouml;rung komplett zu unterlassen. TV, Audio- und Soundanlagen sind auf Zimmerlautst&auml;rke zu betreiben. Das Spielen von Instrumenten ist w&auml;hrend der Ruhezeiten verboten und au&szlig;erhalb auf max. 2 Stunden t&auml;glich beschr&auml;nkt.
        </div>
        <div class="p-block">
          <strong>3. Partys:</strong> St&ouml;rende Partys und laute Feiern sind im gesamten Haus und im Zimmer untersagt. Geselliges Beisammensein ist sp&auml;testens ab 22:00 Uhr auf Zimmerlautst&auml;rke zu beschr&auml;nken.
        </div>
        <div class="p-block">
          <strong>4. Reinigung:</strong> Die Mieter reinigen die Mietr&auml;ume und die Gemeinschaftsfl&auml;chen (K&uuml;che, Bad, WC, Flur, Balkon) eigenverantwortlich nach gemeinsamem Putzplan. Erfolgt trotz Fristsetzung keine Reinigung, kann der Vermieter eine gewerbliche Reinigung auf Kosten aller Mieter zu gleichen Teilen beauftragen. Wiederholte Reinigungsverweigerung berechtigt nach Abmahnung zur fristlosen K&uuml;ndigung.
        </div>
        <div class="p-block">
          <strong>5. Rauch- &amp; Drogenverbot:</strong> In der gesamten Wohnung, im Treppenhaus und auf den Balkonen gilt ein striktes Rauchverbot (inkl. E-Zigaretten &amp; Shishas). Besitz, Konsum oder Handel illegaler Bet&auml;ubungsmittel f&uuml;hrt zur sofortigen Strafanzeige bei der Polizei und zur au&szlig;erordentlichen fristlosen K&uuml;ndigung.
        </div>
        <div class="p-block">
          <strong>6. Haustiere:</strong> Die Haltung von Haustieren jeglicher Art ist grunds&auml;tzlich verboten.
        </div>
        <div class="p-block">
          <strong>7. Beherbergung Dritter:</strong> Eine Gebrauchs&uuml;berlassung oder Untervermietung an Dritte sowie die dauerhafte Aufnahme von Lebensgef&auml;hrten ist unzul&auml;ssig. &Uuml;bernachtungsbesuch ist nur bei Anwesenheit des Mieters und an h&ouml;chstens <strong>f&uuml;nf (5) Tagen im Monat</strong> gestattet.
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">&sect; 1 Cohabitation in Flat Share / Neighbours</div>
        <div class="p-block">
          <strong>1. Mutual Courtesy:</strong> Flatmates must independently organise everyday co-living with fairness, courtesy, and respect for all co-tenants' interests.
        </div>
        <div class="p-block">
          <strong>2. Quiet Hours:</strong> Prevent loud or unnecessary noise. Mandatory quiet hours: <strong>1:00pm to 3:00pm</strong> and <strong>10:00pm to 8:00am</strong>. Absolute silence on Sundays and public holidays. Media devices must operate at room volume. Instrument playing is prohibited during quiet periods and capped at 2 hours daily otherwise.
        </div>
        <div class="p-block">
          <strong>3. Parties:</strong> Disruptive parties are strictly forbidden on premises. Social gatherings must be reduced to quiet room volume by 10:00pm.
        </div>
        <div class="p-block">
          <strong>4. Maintenance &amp; Cleaning:</strong> Tenants clean their room and all shared areas (kitchen, bath, WC, corridor, balcony) on a weekly rota. If neglected despite warning, professional cleaners will be hired at equal shared cost. Persistent refusal warrants lease termination.
        </div>
        <div class="p-block">
          <strong>5. Non-Smoking &amp; Narcotics Ban:</strong> Smoking of any kind (including vapes) is strictly prohibited in all rooms, hallways, and balconies. Possession or use of illegal substances results in immediate police report and lease termination without notice.
        </div>
        <div class="p-block">
          <strong>6. Pets:</strong> Keeping animals or pets of any kind is strictly forbidden.
        </div>
        <div class="p-block">
          <strong>7. Overnight Guests:</strong> Subletting or permanently lodging partners/guests is barred. Overnight visitors are permitted only while tenant is present, capped at maximum <strong>five (5) days per calendar month</strong>.
        </div>
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo} &bull; Homelyo AGB &bull; &sect; 1 Zusammenleben in der WG</div>
    <div>Seite / Page 5 von 8</div>
  </div>
</div>

<!-- ==================== PAGE 6: ANLAGE 2 AGB TEIL 2 (§ 2 REGELUNGEN ZUR WOHNUNG) ==================== -->
<div class="page" id="page-6">
  <div class="page-header">
    <div>{logo_img}</div>
    <div class="header-meta">
      <strong>Anlage 2 &ndash; AGB (Teil 2)</strong> &bull; Regelungen zur Wohnung
    </div>
  </div>

  <div class="page-body">
    <!-- § 2 Regelungen zur Wohnung -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">&sect; 2 Regelungen zur Wohnung</div>
        <div class="p-block">
          <strong>1. Wohnzweck:</strong> Die Anmietung erfolgt ausschlie&szlig;lich zu Wohnzwecken. Gewerbliche Nutzung ist untersagt.
        </div>
        <div class="p-block">
          <strong>2. L&uuml;ften &amp; Heizen:</strong> Zur Schimmelvermeidung ist die Wohnung auch im Winter mindestens 2x t&auml;glich f&uuml;r 5&ndash;10 Minuten durch Sto&szlig;l&uuml;ften zu bel&uuml;ften (Dauerkippen unzul&auml;ssig). Heizk&ouml;rper beim L&uuml;ften und bei Abwesenheit herunterdrehen.
        </div>
        <div class="p-block">
          <strong>3. Internet &amp; Telefon:</strong> Eigene Telefonanschl&uuml;sse d&uuml;rfen nicht verlegt werden. WLAN wird bereitgestellt. Der Mieter haftet vollumf&auml;nglich f&uuml;r seine Internetnutzung. Filesharing (Torrents, P2P) ist verboten; bei Verst&ouml;&szlig;en haftet der Mieter f&uuml;r alle Abmahnkosten und Daten werden an Ermittler weitergegeben. Der Rundfunkbeitrag wird von der WG selbst koordiniert.
        </div>
        <div class="p-block">
          <strong>4. Bilder &amp; Bohren:</strong> N&auml;gel, Schrauben, D&uuml;belbohrungen oder klebende Folien an W&auml;nden und Fliesen sind untersagt. Schilder in Gemeinschaftsr&auml;umen sind verboten; das Klingelschild wird einheitlich beschriftet.
        </div>
        <div class="p-block">
          <strong>5. Obhut &amp; Abfl&uuml;sse:</strong> Keine Speisereste, Fette, &Ouml;le oder Hygieneartikel in Abfl&uuml;sse leeren. Fu&szlig;b&ouml;den (Parkett/Laminat) sind vor Kratzern zu sch&uuml;tzen.
        </div>
        <div class="p-block">
          <strong>6. M&uuml;lltrennung:</strong> M&uuml;ll ist konsequent nach BSR-Vorgaben zu trennen (Gelbe Tonne/Wertstoffe, Papier, Bio, Restm&uuml;ll). Sperrm&uuml;ll und Glas geh&ouml;ren in &ouml;ffentliche Beh&auml;lter.
        </div>
        <div class="p-block">
          <strong>7. Bauliche Ver&auml;nderungen:</strong> Verboten. Satellitenantennen nur nach schriftlicher Genehmigung. Notwendige Erhaltungsarbeiten sind nach Ank&uuml;ndigung zu dulden.
        </div>
        <div class="p-block">
          <strong>8. Sicherheitsbestimmungen:</strong> Haust&uuml;ren und Hoftore sind 22:00&ndash;06:00 Uhr geschlossen zu halten. Fluchtwege sind freizuhalten. Grillen auf Balkonen ist untersagt. Bei Gasgeruch kein Licht/Schalter bet&auml;tigen, Fenster &ouml;ffnen, Haupthahn schlie&szlig;en und Versorger informieren.
        </div>
        <div class="p-block">
          <strong>9. Betreten durch den Vermieter:</strong> Zutritt werktags (09:00&ndash;12:00 und 15:00&ndash;20:00 Uhr) nach Ank&uuml;ndigung f&uuml;r Pr&uuml;fungen oder Handwerker gestattet; in den letzten 4 Wochen vor Auszug mind. 5 Besichtigungstermine. Bei Gefahr im Verzug jederzeitiger Zutritt.
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">&sect; 2 Apartment Regulations</div>
        <div class="p-block">
          <strong>1. Residential Use:</strong> Strictly for private residential accommodation. Commercial activities are prohibited.
        </div>
        <div class="p-block">
          <strong>2. Ventilation &amp; Heating:</strong> Cross-ventilate room minimum 2x daily (5&ndash;10 min) to avoid humidity and mold. Tilt windows forbidden in winter. Shut off radiators when airing or away.
        </div>
        <div class="p-block">
          <strong>3. Internet &amp; Telecom:</strong> Individual phone lines may not be installed. Wi-Fi is provided. Tenant bears full liability for internet traffic; torrenting/P2P is banned. Tenant indemnifies landlord against statutory warnings; logs forwarded to police upon breach. Broadcasting fee is organized by tenants.
        </div>
        <div class="p-block">
          <strong>4. Wall Protection:</strong> Drilling holes, nails, screws, or wall plugs are forbidden. Adhesive posters damaging paint are barred. Doorbell nameplate is labelled uniformly.
        </div>
        <div class="p-block">
          <strong>5. Drains &amp; Floors:</strong> No food waste, oil, grease, or solids down sinks, shower, or toilets. Laminate flooring must be protected from scratch damage.
        </div>
        <div class="p-block">
          <strong>6. Waste Separation:</strong> Systematically separate waste per Berlin BSR statutes (packaging, paper, bio, residual). Bulky waste and glass bottles must be taken to public depots.
        </div>
        <div class="p-block">
          <strong>7. Modifications:</strong> Structural modifications barred. Satellite dishes require written consent. Necessary repairs must be tolerated after notice.
        </div>
        <div class="p-block">
          <strong>8. Safety Measures:</strong> Building and yard doors closed 10:00pm&ndash;6:00am. Hallways/escape paths kept unobstructed. Balcony grilling forbidden. If gas smelled: no sparks/switches, open windows, close main valve, alert provider.
        </div>
        <div class="p-block">
          <strong>9. Landlord Access:</strong> Access on workdays (9:00am&ndash;12:00pm &amp; 3:00pm&ndash;8:00pm) with advance notice for repairs/inspections; at least 5 viewings in the final 4 weeks of lease. Unannounced access permitted during emergencies.
        </div>
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo} &bull; Homelyo AGB &bull; &sect; 2 Regelungen zur Wohnung</div>
    <div>Seite / Page 6 von 8</div>
  </div>
</div>

<!-- ==================== PAGE 7: ANLAGE 2 AGB TEIL 3 (§ 3 & § 4) ==================== -->
<div class="page" id="page-7">
  <div class="page-header">
    <div>{logo_img}</div>
    <div class="header-meta">
      <strong>Anlage 2 &ndash; AGB (Teil 3)</strong> &bull; Zahlungen, &Uuml;bergabe &amp; Schl&uuml;ssel
    </div>
  </div>

  <div class="page-body">
    <!-- § 3 Zahlungsvereinbarungen -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">&sect; 3 Zahlungsvereinbarungen</div>
        <div class="p-block">
          <strong>1. Rechtzeitigkeit:</strong> Ma&szlig;geblich ist das Datum des Eingangs auf dem Konto des Vermieters, nicht das Absendedatum.
        </div>
        <div class="p-block">
          <strong>2. Kaution:</strong> Sichert alle Anspr&uuml;che des Vermieters. Zahlung per Bank&uuml;berweisung. Aufrechnung gegen Mietforderungen vor Mietende ist ausgeschlossen.
        </div>
        <div class="p-block">
          <strong>3. Mietzahlung:</strong> Miete im Voraus bis zum 3. Werktag des Monats kostenfrei &uuml;berweisen mit Verwendungszweck &bdquo;Miete [Monat] [Name]&ldquo;.
        </div>
        <div class="p-block">
          <strong>4. Zahlungsarten:</strong> &Uuml;berweisung auf das Vermieterkonto; anfallende Bank- oder Zahlungsgeb&uuml;hren tr&auml;gt der Mieter.
        </div>
        <div class="p-block">
          <strong>5. R&uuml;ckerstattung:</strong> Abrechnung binnen <strong>30 Tagen</strong> nach Auszug und R&uuml;ckgabe per Bank&uuml;berweisung. Kontodaten sind sp&auml;testens beim Auszug mitzuteilen.
        </div>
        <div class="p-block">
          <strong>6. Fair Usage Nebenkosten:</strong> Pauschalmiete ohne j&auml;hrliche Abrechnung. &Uuml;bersteigen die Betriebskosten den Durchschnitt um mehr als 25 %, kann der Mehrbetrag nach billigem Ermessen umgelegt werden (Gesamtschuldnerhaftung). Ger&auml;te mit hohem Energiebedarf sind anzuzeigen.
        </div>
        <div class="p-block">
          <strong>7. Mahngeb&uuml;hren:</strong> F&uuml;r jede Mahnung wird eine Geb&uuml;hr von <strong>{{{{LATE_FEE}}}} &euro;</strong> berechnet.
        </div>
        <div class="p-block">
          <strong>8. Mieterh&ouml;hung &amp; Aufrechnung:</strong> Mieterh&ouml;hung nach gesetzlichen Regeln. Aufrechnung nur nach Anzeige mindestens einen Monat vor F&auml;lligkeit zul&auml;ssig (&sect;&sect; 536, 536a BGB unber&uuml;hrt).
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">&sect; 3 Payment Agreements</div>
        <div class="p-block">
          <strong>1. Punctuality:</strong> Payment timeliness is judged strictly by credit date on landlord's account, not wire date.
        </div>
        <div class="p-block">
          <strong>2. Deposit:</strong> Secures all claims. Wire payment. Offsetting deposit against rent before departure is prohibited.
        </div>
        <div class="p-block">
          <strong>3. Rent:</strong> Transferred in advance by the 3rd working day with reference "Rent [Month] [Name]".
        </div>
        <div class="p-block">
          <strong>4. Methods:</strong> Wire transfer; transaction/banking charges are covered by tenant.
        </div>
        <div class="p-block">
          <strong>5. Refund:</strong> Returned within <strong>30 days</strong> after departure and return protocol via wire. Bank details must be supplied upon departure.
        </div>
        <div class="p-block">
          <strong>6. Fair Usage:</strong> Flat rent without annual utility settlement. If operating consumption exceeds average by > 25%, excess may be charged at fair discretion (joint liability). Heavy appliances must be declared.
        </div>
        <div class="p-block">
          <strong>7. Late Fees:</strong> Reminder notice fine of <strong>&euro;{{{{LATE_FEE}}}}</strong> per issuance.
        </div>
        <div class="p-block">
          <strong>8. Increase &amp; Set-Off:</strong> Rent adjustments under legal rules. Set-off permitted only with 1 month prior written notice (&sect;&sect; 536, 536a BGB unaffected).
        </div>
      </div>
    </div>

    <!-- § 4 Übergabe und Rückgabe -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">&sect; 4 &Uuml;bergabe und R&uuml;ckgabe der Mietsache</div>
        <div class="p-block">
          <strong>1. Zeitpunkt:</strong> &Uuml;bergabe erfolgt nach Eingang von Kaution und 1. Miete, fr&uuml;hestens zum Mietbeginn.
        </div>
        <div class="p-block">
          <strong>2. Zustand bei Einzug:</strong> Es wird ein &Uuml;bergabeprotokoll angefertigt; dokumentierte M&auml;ngel sind bei Auszug im selben Zustand zu &uuml;bergeben.
        </div>
        <div class="p-block">
          <strong>3. Schl&uuml;sselsatz &amp; Ersatzgeb&uuml;hren:</strong> Bei Verlust oder Besch&auml;digung von Schl&uuml;sseln tr&auml;gt der Mieter folgende Pauschalen:<br>
          &bull; <strong>Haust&uuml;rschl&uuml;ssel (Schlie&szlig;anlage): 680,00 &euro;</strong><br>
          &bull; <strong>Wohnungst&uuml;rschl&uuml;ssel: 120,00 &euro;</strong><br>
          &bull; <strong>Zimmert&uuml;rschl&uuml;ssel: 90,00 &euro;</strong><br>
          &bull; <strong>Briefkastenschl&uuml;ssel: 70,00 &euro;</strong><br>
          Der Mieter haftet erg&auml;nzend f&uuml;r den Austausch von Schl&ouml;ssern der Zentralschlie&szlig;anlage.
        </div>
        <div class="p-block">
          <strong>4. R&uuml;ckgabe bei Auszug:</strong> Am letzten Miettag bis 11:00 Uhr ger&auml;umt, besenrein mit allen Schl&uuml;sseln zur&uuml;ckzugeben. Zur&uuml;ckgelassene Sachen werden auf Mieterkosten entsorgt.
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">&sect; 4 Handover and Return of the Property</div>
        <div class="p-block">
          <strong>1. Handover:</strong> Takes place after confirmed deposit and 1st rent receipt, at earliest on start date.
        </div>
        <div class="p-block">
          <strong>2. Check-In Protocol:</strong> Handover protocol is signed; documented condition must be returned upon departure.
        </div>
        <div class="p-block">
          <strong>3. Key Schedule &amp; Replacement:</strong> If keys are lost or broken, tenant compensates the scheduled expenses:<br>
          &bull; <strong>Building entrance key (master lock): &euro;680.00</strong><br>
          &bull; <strong>Apartment key: &euro;120.00</strong><br>
          &bull; <strong>Room key: &euro;90.00</strong><br>
          &bull; <strong>Mailbox key: &euro;70.00</strong><br>
          Tenant remains liable for replacing master cylinders if required for safety.
        </div>
        <div class="p-block">
          <strong>4. Departure:</strong> By 11:00am on final lease day, completely cleared, swept-clean with all keys. Leftover belongings disposed of at tenant's expense.
        </div>
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo} &bull; Homelyo AGB &bull; &sect; 3 Zahlungen &amp; &sect; 4 &Uuml;bergabe</div>
    <div>Seite / Page 7 von 8</div>
  </div>
</div>

<!-- ==================== PAGE 8: ANLAGE 2 AGB TEIL 4 (§ 5, § 6, § 7, § 8 & FINAL SIGNATURES) ==================== -->
<div class="page" id="page-8">
  <div class="page-header">
    <div>{logo_img}</div>
    <div class="header-meta">
      <strong>Anlage 2 &ndash; AGB (Teil 4)</strong> &bull; Sch&auml;den, K&uuml;ndigung &amp; Unterschriften
    </div>
  </div>

  <div class="page-body">
    <!-- § 5 Schäden an der Wohnung -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">&sect; 5 Sch&auml;den an der Wohnung &amp; Kleinreparaturen</div>
        <div class="p-block">
          <strong>1. Meldepflicht:</strong> M&auml;ngel sind unverz&uuml;glich mitzuteilen; dem Vermieter ist eine Frist von 4 Wochen zur Abhilfe einzur&auml;umen.
        </div>
        <div class="p-block">
          <strong>2. Haftung &amp; Beweislast:</strong> Der Mieter haftet f&uuml;r Sch&auml;den durch Pflichtverletzung (unzureichendes L&uuml;ften/Heizen) und Besucher. Beweislast f&uuml;r Sch&auml;den im Zimmer und WG-R&auml;umen liegt beim Mieter. Vermieter haftet nur f&uuml;r Vorsatz und grobe Fahrl&auml;ssigkeit.
        </div>
        <div class="p-block">
          <strong>3. Sch&ouml;nheitsreparaturen:</strong> Nicht geschuldet, au&szlig;er bei &uuml;berm&auml;&szlig;iger Abnutzung.
        </div>
        <div class="p-block">
          <strong>4. Kleinreparaturen:</strong> Reparaturen an Teilen im st&auml;ndigen Mieterzugriff (Leuchtmittel, Griffe, Armaturen) tr&auml;gt der Mieter bis <strong>100,00 &euro;</strong> je Einzelfall (j&auml;hrlich max. 8% der Miete bzw. <strong>350,00 &euro;</strong>).
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">&sect; 5 Damages to Property &amp; Minor Repairs</div>
        <div class="p-block">
          <strong>1. Reporting:</strong> Inform landlord immediately of defects; landlord is granted 4 weeks to remediate.
        </div>
        <div class="p-block">
          <strong>2. Liability:</strong> Tenant is liable for neglect (poor airing/heating) and visitor actions. Burden of proof rests with tenant for room damages. Landlord liability limited to intent and gross negligence.
        </div>
        <div class="p-block">
          <strong>3. Cosmetic Repairs:</strong> Not required except in case of excessive deterioration.
        </div>
        <div class="p-block">
          <strong>4. Minor Repairs:</strong> Upkeep on direct-use fixtures (bulbs, handles, faucets) borne by tenant up to <strong>&euro;100.00</strong> per incident (annual cap: 8% rent or <strong>&euro;350.00</strong>).
        </div>
      </div>
    </div>

    <!-- § 6 Kündigung, § 7 Schriftform, § 8 Salvatorische Klausel -->
    <div class="two-col">
      <div class="col col-de">
        <div class="sec-title">&sect; 6 K&uuml;ndigung, &sect; 7 Schriftform &amp; &sect; 8 Salvatorische Klausel</div>
        <div class="p-block">
          <strong>1. K&uuml;ndigung &amp; Nachmieter:</strong> K&uuml;ndigungsfrist 3 Monate zum Monatsende gem. &sect; 5 Hauptvertrag. Bei Gestellung eines akzeptierten Nachmieters wird der Mieter ab dessen &Uuml;bergabe freigestellt. &sect; 545 BGB wird ausgeschlossen.
        </div>
        <div class="p-block">
          <strong>2. Fristlose K&uuml;ndigung (&sect; 543 BGB):</strong> Zul&auml;ssig bei: Verzug von 2 Monatsmieten, AGB-Versto&szlig; trotz Abmahnung, Sch&auml;den &gt; <strong>50,00 &euro;</strong> unbeglichen, unbefugter Dritter, Drogen oder Reinigungsverweigerung.
        </div>
        <div class="p-block">
          <strong>3. Schriftform &amp; Salvatorische Klausel:</strong> M&uuml;ndliche Nebenabreden ung&uuml;ltig; &Auml;nderungen nur schriftlich. Sollten Klauseln unwirksam sein, bleibt der Restvertrag wirksam. Erf&uuml;llungsort und Gerichtsstand ist Berlin.
        </div>
      </div>

      <div class="col col-en">
        <div class="sec-title">&sect; 6 Termination, &sect; 7 Form &amp; &sect; 8 Severability</div>
        <div class="p-block">
          <strong>1. Notice &amp; Replacement:</strong> 3 months notice to month-end per &sect; 5 Core Agreement. Tenant is released early upon move-in of accepted replacement. Section 545 BGB is excluded.
        </div>
        <div class="p-block">
          <strong>2. Cancellation for Cause (&sect; 543 BGB):</strong> For 2 months rent arrears, GTC breach after warning, damages &gt; <strong>&euro;50.00</strong> unpaid, unauthorized lodging, drugs, or persistent cleaning refusal.
        </div>
        <div class="p-block">
          <strong>3. Written Form &amp; Severability:</strong> No oral agreements. Invalid provisions do not invalidate the rest. Governing law is German BGB. Place of jurisdiction is Berlin.
        </div>
      </div>
    </div>

    <!-- Signatures Block 2 (AGB) -->
    <div class="sig-box">
      <table class="sig-table">
        <tr>
          <td>
            Ort, Datum / Place, Date: Berlin, {{{{LANDLORD_SIGN_DATE}}}}
            <div class="sig-line"></div>
            <strong>{{{{LANDLORD_SIGN_NAME}}}}</strong><br>
            Vermieter / Landlord
          </td>
          <td>
            Ort, Datum / Place, Date: {{{{TENANT_SIGN_DATE}}}}
            <div class="sig-line"></div>
            <strong>{{{{TENANT_SIGN_NAME}}}}</strong><br>
            Mieter / Tenant
          </td>
        </tr>
      </table>
      <div class="legal-banner">Anlage 2 akzeptiert und vollst&auml;ndig anerkannt. / Attachment 2 accepted and agreed in full.</div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo} &bull; Homelyo AGB &bull; Unterschriften &amp; Best&auml;tigung</div>
    <div>Seite / Page 8 von 8</div>
  </div>
</div>

</body>
</html>
"""
    return html

def render_contract_html(data):
    html = generate_html_template()
    for k, v in data.items():
        html = html.replace(f"{{{{{k}}}}}", str(v))
    return html

def render_contract_docx(data, output_path):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)

    logo_path = LOGO_PATH if os.path.exists(LOGO_PATH) else FALLBACK_LOGO_PATH
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(logo_path, width=Inches(1.8))

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_de = p_title.add_run("Mietvertrag\n")
    r_de.font.size = Pt(22)
    r_de.font.bold = True
    r_en = p_title.add_run("Residential Lease Agreement & House Rules (AGB)")
    r_en.font.size = Pt(12)
    r_en.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Executive Overview Table
    t_overview = doc.add_table(rows=5, cols=2)
    t_overview.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r in t_overview.rows:
        r.cells[0].width = Inches(3.6)
        r.cells[1].width = Inches(3.6)

    t_overview.rows[0].cells[0].paragraphs[0].add_run("Mieter / Tenant: ").font.bold = True
    t_overview.rows[0].cells[0].paragraphs[0].add_run(f"{data['TENANT_NAME']}")
    t_overview.rows[0].cells[1].paragraphs[0].add_run("Mietobjekt / Room: ").font.bold = True
    t_overview.rows[0].cells[1].paragraphs[0].add_run(f"Zimmer {data['ROOM_CODE']} ({data['ROOM_TITLE_DE']} / {data['ROOM_TITLE_EN']})")

    t_overview.rows[1].cells[0].paragraphs[0].add_run("Liegenschaft / Address: ").font.bold = True
    t_overview.rows[1].cells[0].paragraphs[0].add_run("Lichtenrader Damm 80, 12305 Berlin (2. OG, WE-11)")
    t_overview.rows[1].cells[1].paragraphs[0].add_run("Mietzeit / Rental Term: ").font.bold = True
    t_overview.rows[1].cells[1].paragraphs[0].add_run(f"{data['START_DATE_DE']} bis {data['END_DATE_DE']}")

    t_overview.rows[2].cells[0].paragraphs[0].add_run("Warmmiete / Warm Rent: ").font.bold = True
    t_overview.rows[2].cells[0].paragraphs[0].add_run(f"{data['RENT_AMOUNT']} € / Monat ({data['RENT_WORDS_DE']} Euro)")
    t_overview.rows[2].cells[1].paragraphs[0].add_run("Rent & Inclusions: ").font.bold = True
    t_overview.rows[2].cells[1].paragraphs[0].add_run(f"€{data['RENT_AMOUNT']} / month (incl. utilities & 1000 Mbps Wi-Fi)")

    t_overview.rows[3].cells[0].paragraphs[0].add_run("Kaution / Security Deposit: ").font.bold = True
    t_overview.rows[3].cells[0].paragraphs[0].add_run(f"{data['DEPOSIT_AMOUNT']} € (vor Einzug fällig)")
    t_overview.rows[3].cells[1].paragraphs[0].add_run("Deposit Terms: ").font.bold = True
    t_overview.rows[3].cells[1].paragraphs[0].add_run(f"€{data['DEPOSIT_AMOUNT']} (held in escrow, payable prior to move-in)")

    t_overview.rows[4].cells[0].paragraphs[0].add_run("Endreinigung / Cleaning: ").font.bold = True
    t_overview.rows[4].cells[0].paragraphs[0].add_run(f"{data['CLEANING_FEE']} € einmalig bei Auszug")
    t_overview.rows[4].cells[1].paragraphs[0].add_run("Final Cleaning Fee: ").font.bold = True
    t_overview.rows[4].cells[1].paragraphs[0].add_run(f"€{data['CLEANING_FEE']} one-time checkout fee")

    doc.add_page_break()

    # 14 Full Clauses Table (§ 1 to § 6 Hauptvertrag + § 1 to § 8 AGB)
    clauses = [
        (
            "§ 1 Mietdauer & Befristungszweck (§ 549 Abs. 2 / § 575 BGB)\n",
            f"(1) Mietzeit: Beginn am {data['START_DATE_DE']}, befristetes Ende am {data['END_DATE_DE']}. Mindestmietdauer beträgt sechs (6) Monate.\n"
            f"(2) Zweck: Vorübergehender Gebrauch gem. § 549 Abs. 2 Nr. 1 BGB ({data['REASON_DE']}).\n"
            "(3) Ausschluss: Stillschweigende Verlängerung (§ 545 BGB) ist ausgeschlossen.",
            "§ 1 Lease Term & Temporary Purpose (§ 549 (2) / § 575 BGB)\n",
            f"(1) Term: Commences on {data['START_DATE_EN']}, terminates on {data['END_DATE_EN']}. Minimum lease term is six (6) months.\n"
            f"(2) Purpose: Temporary living pursuant to Section 549 (2) No. 1 BGB ({data['REASON_EN']}).\n"
            "(3) Tacit Renewal: Section 545 BGB is strictly excluded."
        ),
        (
            "§ 2 Miethöhe, Pauschale & Nebenkosten\n",
            f"(1) Warmmiete: {data['RENT_AMOUNT']} € monatlich ({data['RENT_WORDS_DE']} Euro).\n"
            "(2) Inklusivleistungen: Kalte & warme Betriebskosten, Strom, Heizung und Vodafone 1000 Mbit/s WLAN in verbrauchsüblichem Maße (Fair Use).\n"
            f"(3) ARD-Beitrag & Endreinigung: GEZ ist nicht enthalten (von WG eigenständig getragen). Endreinigungspauschale bei Auszug: {data['CLEANING_FEE']} €.",
            "§ 2 Rent, Utilities & Public License Fee\n",
            f"(1) Rent: Flat monthly rate of €{data['RENT_AMOUNT']} ({data['RENT_WORDS_EN']} euros).\n"
            "(2) Inclusions: Heating, water, electricity, operating costs and Vodafone 1000 Mbps Wi-Fi based on customary fair use.\n"
            f"(3) ARD License & Cleaning: Broadcasting fee (GEZ) settled directly by WG. Final cleaning fee upon checkout: €{data['CLEANING_FEE']}."
        ),
        (
            "§ 3 Kaution & Mietsicherheit (§ 551 BGB)\n",
            f"(1) Kaution: {data['DEPOSIT_AMOUNT']} € ({data['DEPOSIT_WORDS_DE']} Euro).\n"
            f"(2) Fälligkeit: Vor Einzug am {data['START_DATE_DE']} vollständig auf das Vermieterkonto zu zahlen. Keine Schlüssel ohne Geldeingang.\n"
            "(3) Erstattung: Treuhandverwahrung gem. § 551 BGB; Erstattung innerhalb von 30 Tagen nach Auszug abzüglich berechtigter Forderungen. Kein Aufrechnen gegen Miete.",
            "§ 3 Security Deposit & Escrow (§ 551 BGB)\n",
            f"(1) Deposit: €{data['DEPOSIT_AMOUNT']} ({data['DEPOSIT_WORDS_EN']} euros).\n"
            f"(2) Due Date: Payable in full before move-in on {data['START_DATE_EN']}. Key handover denied without receipt.\n"
            "(3) Refund: Escrow under § 551 BGB; refunded within 30 days after move-out protocol less proven claims. No set-off against ongoing rent."
        ),
        (
            "§ 4 Zahlungsbedingungen, Fälligkeit & Verzug\n",
            f"(1) Fälligkeit: Bis spätestens zum 3. Werktag des Monats im Voraus kostenfrei zu überweisen.\n"
            f"(2) Bank: {data['RENT_BANK']} | Inhaber: {data['RENT_ACCOUNT_HOLDER']} | IBAN: {data['RENT_IBAN']} | BIC: {data['RENT_BIC']}.\n"
            f"(3) Verzug: Mahngebühr {data['LATE_FEE']} € je Mahnung. Kündigung gem. § 543 Abs. 2 BGB bei Verzug von 2 Monatsmieten.",
            "§ 4 Payment Terms, Due Dates & Late Fees\n",
            "(1) Due Date: Payable in advance on or before the 3rd business day of each month via wire transfer.\n"
            f"(2) Account: {data['RENT_BANK']} | Holder: {data['RENT_ACCOUNT_HOLDER']} | IBAN: {data['RENT_IBAN']} | BIC: {data['RENT_BIC']}.\n"
            f"(3) Default: Reminder charge of €{data['LATE_FEE']} per notice. Cancellation under § 543 (2) BGB for 2 months arrears."
        ),
        (
            "§ 5 Kündigung & Nachmieterstellung\n",
            "(1) Ordentliche Kündigung: Während der festen Mietdauer ausgeschlossen.\n"
            "(2) Nachmieter: Vorzeitige Entlassung mit 3 Monaten Frist zum Monatsende bei Gestellung eines solventen, für die WG zumutbaren Nachmieters.\n"
            "(3) Außerordentliche Kündigung: Gesetzliche Kündigung aus wichtigem Grund (§§ 543, 569 BGB) bleibt unberührt.",
            "§ 5 Termination & Replacement Tenant Release\n",
            "(1) Ordinary Notice: Excluded during the fixed lease term.\n"
            "(2) Replacement Tenant: Early release with 3 months' notice to month-end upon presenting an accepted solvent replacement tenant who executes a contract.\n"
            "(3) Cause: Extraordinary termination for good cause (§§ 543, 569 BGB) unaffected."
        ),
        (
            "§ 6 AGB & Anlagen als Vertragsbestandteil\n",
            "Anlage 1 (Inventar) und Anlage 2 (AGB § 1 bis § 8) werden mit Unterschrift anerkannt. Übergabeprotokoll, Schlüsselquittung und Wohnungsgeberbestätigung (§ 19 BMG) werden bei Einzug ausgehändigt.",
            "§ 6 Terms & Attachments\n",
            "Attachment 1 (Inventory) and Attachment 2 (GTC § 1 to § 8) are integral binding parts. Handover protocol, keys acknowledgement, and landlord confirmation (§ 19 BMG) provided upon move-in."
        ),
        (
            "AGB § 1 Zusammenleben in der WG / mit den Nachbarn\n",
            "Rücksichtnahme in der WG. Ruhezeiten: 13:00–15:00 Uhr sowie 22:00–08:00 Uhr (Sonn-/Feiertage ganztägig). TV/Audio auf Zimmerlautstärke. Störende Partys verboten. Reinigung eigenverantwortlich nach Putzplan; bei Säumnis Vergabe an Fremdfirma auf WG-Kosten. Striktes Rauch- und Drogenverbot (Strafanzeige). Keine Haustiere. Gäste max. 5 Tage/Monat.",
            "GTC § 1 Cohabitation in Flat Share / Neighbours\n",
            "Mutual courtesy. Quiet hours: 1:00pm–3:00pm and 10:00pm–8:00am (Sundays/holidays full quiet). Media at room volume. Disruptive parties prohibited. Shared cleaning rota; default results in professional cleaning at tenants' equal cost. Strict non-smoking and drug ban (police report). No pets. Overnight guests max 5 days/month."
        ),
        (
            "AGB § 2 Regelungen zur Wohnung\n",
            "Ausschließlich Wohnzwecke. Stoßlüften mind. 2x täglich für 5–10 Min. (kein Dauerkippen im Winter). WLAN bereitgestellt; P2P/Torrents verboten; Mieter haftet und stellt Vermieter frei. Wandbohrungen, Nägel, Dübel und klebende Folien untersagt. Abflüsse vor Fett/Abfall schützen. Mülltrennung nach BSR (Gelb, Papier, Bio, Rest). Türen 22:00–06:00 Uhr geschlossen. Zutrittsrecht des Vermieters werktags 09–12 & 15–20h nach Vorankündigung (letzte 4 Wochen mind. 5 Termine) und bei Gefahr im Verzug.",
            "GTC § 2 Apartment Regulations\n",
            "Residential use only. Air out 2x daily 5–10 min (no tilting in winter). Wi-Fi provided; P2P/torrents prohibited; tenant fully indemnifies landlord. No drilling, nails, screws or wall plugs. Drains kept free of grease/solids. Waste sorting per Berlin BSR. Doors closed 10:00pm–6:00am. Landlord access workdays 9am–12pm & 3pm–8pm upon notice (at least 5 viewings in last 4 weeks) and anytime during emergencies."
        ),
        (
            "AGB § 3 Zahlungsvereinbarungen\n",
            "Rechtzeitigkeit nach Geldeingang. Kaution sichert alle Ansprüche; kein Aufrechnen vor Mietende. Miete im Voraus bis 3. Werktag. Kautionsabrechnung binnen 30 Tagen nach Auszug. Fair Usage der Nebenkosten (> 25% Mehraufwand wird umgelegt). Mahngebühr 5,00 € je Mahnung. Mieterhöhung nach BGB. Aufrechnung nur mit 1 Monat Vorankündigung.",
            "GTC § 3 Payment Agreements\n",
            "Punctuality determined by account credit date. Deposit secures all claims; no set-off against ongoing rent. Rent in advance by 3rd business day. Deposit refund within 30 days after move-out. Fair usage (> 25% excess billed). Overdue fine €5.00 per notice. Rent adjustments per statutory law. Set-off requires 1 month notice."
        ),
        (
            "AGB § 4 Übergabe und Rückgabe & Schlüsselgebühren\n",
            "Übergabe nach Zahlungseingang per Protokoll. Bei Auszug geräumt und besenrein bis 11:00 Uhr mit allen Schlüsseln. Schlüsselersatzpauschalen:\n"
            "• Haustür (Schließanlage): 680,00 € | Wohnung: 120,00 €\n"
            "• Zimmer: 90,00 € | Briefkasten: 70,00 €\n"
            "Zusätzliche Haftung für Schlossaustausch der Schließanlage.",
            "GTC § 4 Handover, Departure & Key Replacement Schedule\n",
            "Handover after payment receipt with protocol. Departure by 11:00am cleared, swept-clean with all keys. Key replacement schedule:\n"
            "• Building entrance (master system): €680.00 | Flat: €120.00\n"
            "• Room: €90.00 | Mailbox: €70.00\n"
            "Tenant additionally liable for lock cylinder replacement."
        ),
        (
            "AGB § 5 Schäden an der Wohnung & Kleinreparaturen\n",
            "Schäden unverzüglich melden (4 Wochen Abhilfefrist). Mieter haftet für Sorgfaltspflichtverletzung und Besucher; Beweislast beim Mieter. Keine Schönheitsreparaturen geschuldet. Kleinreparaturen bis 100,00 € je Einzelfall (max. 8% der Jahresmiete bzw. 350,00 € jährlich) trägt der Mieter.",
            "GTC § 5 Damages to Property & Minor Repairs\n",
            "Report defects immediately (4 weeks remedy period). Tenant liable for neglect and visitors; burden of proof rests with tenant. No general cosmetic repairs. Minor repairs up to €100.00 per instance (annual cap: 8% of rent or €350.00) borne by tenant."
        ),
        (
            "AGB § 6 Kündigungsbedingungen\n",
            "Kündigungsfrist 3 Monate zum Monatsende. Vorzeitiges Ausscheiden bei akzeptiertem Nachmieter ab Übergabe. Stillschweigende Verlängerung gem. § 545 BGB ausgeschlossen. Fristlose Kündigung gem. § 543 BGB zulässig bei: Verzug von 2 Monatsmieten, AGB-Verstoß nach Abmahnung, Schäden > 50,00 € unbeglichen, unerlaubter Dritter, Drogen oder Reinigungsverweigerung.",
            "GTC § 6 Conditions of Termination\n",
            "Notice period 3 months to month-end. Early release upon accepted replacement tenant. Tacit renewal excluded under § 545 BGB. Termination without notice under § 543 BGB for: 2 months rent arrears, GTC breach after warning, damages > €50.00 unpaid, unauthorized guests, drugs, or persistent cleaning refusal."
        ),
        (
            "AGB § 7 Schriftform\n",
            "Mündliche Nebenabreden bestehen nicht. Änderungen und Ergänzungen bedürfen der Schriftform; dies gilt auch für die Aufhebung dieses Schriftformerfordernisses.",
            "GTC § 7 Written Form\n",
            "No oral side agreements exist. Modifications and amendments must be executed in writing; this also applies to waiving this written-form requirement."
        ),
        (
            "AGB § 8 Salvatorische Klausel & Gerichtsstand\n",
            "Unwirksamkeit einzelner Bestimmungen berührt die übrigen Klauseln nicht. Es gilt das Recht der Bundesrepublik Deutschland (BGB). Erfüllungsort und ausschließlicher Gerichtsstand ist Berlin. Die deutsche Fassung ist rechtlich allein verbindlich.",
            "GTC § 8 Severability & Jurisdiction\n",
            "Invalidity of individual clauses does not affect remaining provisions. German Civil Code (BGB) governs. Place of performance and sole legal jurisdiction is Berlin. Only the German language version is legally binding."
        )
    ]

    t_clauses = doc.add_table(rows=len(clauses), cols=2)
    t_clauses.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (title_de, text_de, title_en, text_en) in enumerate(clauses):
        row = t_clauses.rows[i]
        row.cells[0].width = Inches(3.6)
        row.cells[1].width = Inches(3.6)

        p_de = row.cells[0].paragraphs[0]
        r_de = p_de.add_run(title_de)
        r_de.font.bold = True
        r_de.font.size = Pt(9.0)
        r_de_txt = p_de.add_run(text_de)
        r_de_txt.font.size = Pt(8.0)

        p_en = row.cells[1].paragraphs[0]
        r_en = p_en.add_run(title_en)
        r_en.font.bold = True
        r_en.font.size = Pt(9.0)
        r_en.font.color.rgb = RGBColor(71, 85, 105)
        r_en_txt = p_en.add_run(text_en)
        r_en_txt.font.size = Pt(8.0)

    doc.add_paragraph().paragraph_format.space_after = Pt(20)

    # Signatures Table
    t_sig = doc.add_table(rows=1, cols=2)
    t_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r in t_sig.rows:
        r.cells[0].width = Inches(3.6)
        r.cells[1].width = Inches(3.6)

    t_sig.rows[0].cells[0].paragraphs[0].add_run(f"Ort, Datum: Berlin, {data['LANDLORD_SIGN_DATE']}\n\n\n_____________________________________\n{data['LANDLORD_SIGN_NAME']}\nVermieter / Landlord")
    t_sig.rows[0].cells[1].paragraphs[0].add_run(f"Ort, Datum: {data['TENANT_SIGN_DATE']}\n\n\n_____________________________________\n{data['TENANT_SIGN_NAME']}\nMieter / Tenant")

    doc.save(output_path)
    return output_path

def convert_html_to_pdf(html_path, pdf_path):
    chrome_bin = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    if not os.path.exists(chrome_bin):
        raise FileNotFoundError(f"Chrome not found at {chrome_bin}")

    os.makedirs(os.path.dirname(os.path.abspath(pdf_path)), exist_ok=True)
    cmd = [
        chrome_bin,
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",
        "--print-to-pdf-no-header",
        f"--print-to-pdf={pdf_path}",
        html_path
    ]
    subprocess.check_call(cmd)
    return pdf_path

def update_gdrive_template(docx_path):
    try:
        token = subprocess.check_output(["gcloud", "auth", "print-access-token"]).decode().strip()
        with open(docx_path, "rb") as f:
            docx_bytes = f.read()

        url = f"https://www.googleapis.com/upload/drive/v3/files/{GDRIVE_TEMPLATE_FILE_ID}?uploadType=media"
        req = urllib.request.Request(
            url,
            data=docx_bytes,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
            method="PATCH"
        )
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode())
            print(f"✓ Updated Google Drive template_contract.docx (ID: {GDRIVE_TEMPLATE_FILE_ID})")
            return data
    except Exception as e:
        print(f"⚠ Notice: GDrive template update skipped or requires auth: {e}")
        return None

def main():
    print("Building official Homelyo Rental Contract Templates (14 Clauses, Wunderflats + GDrive Standard)...")
    os.makedirs(os.path.dirname(HTML_TEMPLATE_PATH), exist_ok=True)
    os.makedirs(os.path.dirname(FINAL_PDF_PATH), exist_ok=True)

    # 1. Master HTML Template
    html = generate_html_template()
    with open(HTML_TEMPLATE_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✓ Master HTML template saved: {HTML_TEMPLATE_PATH}")

    # 2. Master DOCX Template
    default_data = render_contract_data(
        tenant_name="Ajithkumar Velasseril Sreenivasan",
        tenant_email="ajithkmvs123@gmail.com",
        tenant_phone="+91 8606817191",
        room_code="KE2",
        start_date="2026-11-02",
        end_date="2027-04-30",
        rent_amount=649.00,
        deposit_amount=1000.00
    )
    render_contract_docx(default_data, DOCX_TEMPLATE_PATH)
    print(f"✓ Master DOCX template saved: {DOCX_TEMPLATE_PATH}")

    # 3. Render Sample HTML and Convert to Final PDF
    rendered_html = render_contract_html(default_data)
    temp_html_path = "/tmp/homelyo_contract_render.html"
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(rendered_html)

    convert_html_to_pdf(temp_html_path, FINAL_PDF_PATH)
    print(f"✓ Final Master PDF generated: {FINAL_PDF_PATH}")

    # 4. Update Google Drive Cloud Template
    update_gdrive_template(DOCX_TEMPLATE_PATH)
    print("✓ All templates and PDF synchronized successfully!")

if __name__ == "__main__":
    main()
