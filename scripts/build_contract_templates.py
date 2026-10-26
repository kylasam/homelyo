#!/usr/bin/env python3
"""
Homelyo Rental Contract Template Builder
----------------------------------------
Builds bilingual (German/English) contracts and complete AGB for Homelyo.
Includes:
- Full Premises Breakdown Table (Street, No, Floor, Flat, Room, Zip, Country, Shared amenities)
- Custom clauses: Advance payment before move-in, Rent due before 4th + €5 late fee
- Cleaning fee: €70.00
- Deposit: €1,000.00
- Key replacement fees: Haustür 680 €, Wohnung 120 €, Zimmer 90 €, Briefkasten 70 €
- ARD / Rundfunkbeitrag coordinated & paid by tenants
- Internet responsibility & full indemnity (P2P/torrents prohibited, immediate termination)
- Complete AGB (§ 1 to § 8) in dual-column format
- Clean signature fields without GMT timestamps
"""

import os
import sys
import datetime
import base64
import subprocess
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

LOGO_PATH = "/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/assistant/assets/Homelyo-logo-header.png"
FALLBACK_LOGO_PATH = "/Users/jayashreesundarasamy/Downloads/Homelyo-logo.png"
HTML_TEMPLATE_PATH = "/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/templates/contract_template.html"
DOCX_TEMPLATE_PATH = "/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/template_contract.docx"

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

    return f"{day} {de_months[m_idx]}, {year}", f"{day} {en_months[m_idx]}, {year}"

def generate_html_template():
    logo_b64 = get_logo_base64()
    logo_img_tag = f'<img src="data:image/png;base64,{logo_b64}" alt="Homelyo Logo" style="height: 38px; width: auto; max-width: 140px; display: inline-block;">' if logo_b64 else '<div style="font-size: 16pt; font-weight: bold; color: #0f172a; letter-spacing: 1px;">HOMELYO</div>'
    footer_logo_tag = f'<img src="data:image/png;base64,{logo_b64}" alt="Homelyo" style="height: 14px; width: auto; opacity: 0.85;">' if logo_b64 else '<span style="font-weight: 700; font-size: 8pt; color: #475569;">homelyo</span>'

    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>Mietvertrag / Rental Agreement - Homelyo</title>
<style>
  @page {{
    size: A4 portrait;
    margin: 11mm 13mm 11mm 13mm;
  }}
  * {{
    box-sizing: border-box;
    -webkit-font-smoothing: antialiased;
  }}
  body {{
    margin: 0;
    padding: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    color: #111827;
    background-color: #ffffff;
    font-size: 8.5pt;
    line-height: 1.34;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }}
  .page {{
    width: 100%;
    page-break-after: always;
    position: relative;
    padding-bottom: 12px;
  }}
  .page:last-child {{
    page-break-after: avoid;
  }}
  .two-col {{
    display: flex;
    flex-direction: row;
    justify-content: space-between;
    width: 100%;
    gap: 18px;
  }}
  .col {{
    flex: 1;
    width: 50%;
  }}
  .page-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 1px solid #e2e8f0;
    padding-bottom: 8px;
    margin-bottom: 12px;
  }}
  .header-left {{
    font-size: 8pt;
    color: #64748b;
    font-weight: 500;
  }}
  .header-right {{
    text-align: right;
  }}
  .page-footer {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-top: 1px solid #e2e8f0;
    padding-top: 8px;
    margin-top: 15px;
    font-size: 7.5pt;
    color: #64748b;
  }}
  .main-title {{
    text-align: center;
    margin: 15px 0 18px 0;
  }}
  .main-title h1 {{
    font-size: 20pt;
    font-weight: 800;
    margin: 0 0 2px 0;
    color: #0f172a;
    letter-spacing: -0.5px;
  }}
  .main-title h2 {{
    font-size: 13pt;
    font-weight: 600;
    color: #475569;
    margin: 0;
  }}
  .party-header {{
    font-size: 10.0pt;
    color: #0f172a;
    margin-bottom: 5px;
    border-bottom: 1.5px solid #cbd5e1;
    padding-bottom: 3px;
  }}
  .field-table {{
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 12px;
  }}
  .field-table td {{
    padding: 3px 0;
    font-size: 8.5pt;
    vertical-align: top;
  }}
  .field-label {{
    width: 38%;
    color: #64748b;
    font-weight: 600;
  }}
  .field-value {{
    width: 62%;
    color: #0f172a;
    font-weight: 500;
  }}
  .premises-box {{
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    padding: 10px 14px;
    margin: 12px 0;
  }}
  .premises-title {{
    font-weight: 700;
    font-size: 9.5pt;
    color: #0f172a;
    margin-bottom: 6px;
  }}
  .premises-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px 14px;
    font-size: 8.5pt;
  }}
  .sec-title {{
    font-size: 9.8pt;
    font-weight: 700;
    color: #0f172a;
    margin-top: 10px;
    margin-bottom: 4px;
    line-height: 1.25;
    border-left: 3px solid #3b82f6;
    padding-left: 6px;
  }}
  .p-block {{
    margin-bottom: 5px;
    text-align: justify;
    text-justify: inter-word;
    font-size: 8.3pt;
  }}
  .highlight-badge {{
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    color: #1e40af;
    padding: 1px 5px;
    border-radius: 3px;
    font-weight: 600;
    display: inline-block;
  }}
  .bank-box {{
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 4px;
    padding: 7px 10px;
    margin: 6px 0;
    font-size: 8.2pt;
  }}
  .bank-row {{
    display: flex;
    justify-content: space-between;
    padding: 1.5px 0;
  }}
  .bank-label {{
    color: #64748b;
    font-weight: 600;
  }}
  .bank-val {{
    font-weight: 600;
    color: #0f172a;
    font-family: monospace;
  }}
  .agb-header {{
    text-align: center;
    background: #0f172a;
    color: #ffffff;
    padding: 8px;
    border-radius: 4px;
    margin: 10px 0 12px 0;
  }}
  .agb-header h3 {{
    margin: 0;
    font-size: 11pt;
    font-weight: 700;
  }}
  .agb-header h4 {{
    margin: 2px 0 0 0;
    font-size: 8.5pt;
    font-weight: 400;
    opacity: 0.9;
  }}
  .sig-block {{
    margin-top: 25px;
    padding-top: 10px;
  }}
  .sig-line {{
    border-top: 1.5px solid #0f172a;
    margin-top: 40px;
    margin-bottom: 4px;
  }}
  .sig-title {{
    font-weight: 700;
    font-size: 8.5pt;
    color: #0f172a;
  }}
  .sig-date {{
    font-size: 8pt;
    color: #475569;
    margin-top: 3px;
  }}
</style>
</head>
<body>

<!-- ==================== PART 1: MIETVERTRAG / LEASE AGREEMENT ==================== -->

<!-- PAGE 1: Parties & Premises -->
<div class="page" id="page-1">
  <div class="page-header">
    <div class="header-left">{logo_img_tag}</div>
    <div class="header-right" style="font-size: 8pt; color: #64748b;">
      <strong>Homelyo Co-Living</strong> &bull; Lichtenrader Damm 80, 12305 Berlin
    </div>
  </div>

  <div class="main-title">
    <h1>Mietvertrag</h1>
    <h2>Residential Lease Agreement</h2>
  </div>

  <!-- Parties -->
  <div class="two-col">
    <!-- Landlord -->
    <div class="col">
      <div class="party-header">Zwischen <strong>Vermieter:</strong></div>
      <table class="field-table">
        <tr>
          <td class="field-label">Name:</td>
          <td class="field-value">{{{{LANDLORD_SIGN_NAME}}}}</td>
        </tr>
        <tr>
          <td class="field-label">Anschrift:</td>
          <td class="field-value">{{{{LANDLORD_ADDRESS_DE}}}}</td>
        </tr>
        <tr>
          <td class="field-label">Kontakt:</td>
          <td class="field-value">support@homelyo.eu | www.homelyo.eu</td>
        </tr>
      </table>
    </div>
    <div class="col">
      <div class="party-header">The <strong>landlord:</strong></div>
      <table class="field-table">
        <tr>
          <td class="field-label">Name:</td>
          <td class="field-value">{{{{LANDLORD_SIGN_NAME}}}}</td>
        </tr>
        <tr>
          <td class="field-label">Address:</td>
          <td class="field-value">{{{{LANDLORD_ADDRESS_EN}}}}</td>
        </tr>
        <tr>
          <td class="field-label">Contact:</td>
          <td class="field-value">support@homelyo.eu | www.homelyo.eu</td>
        </tr>
      </table>
    </div>
  </div>

  <!-- Tenant -->
  <div class="two-col">
    <div class="col">
      <div class="party-header">Und <strong>Mieter:</strong></div>
      <table class="field-table">
        <tr>
          <td class="field-label">Name:</td>
          <td class="field-value"><strong>{{{{TENANT_NAME}}}}</strong></td>
        </tr>
        <tr>
          <td class="field-label">Anschrift:</td>
          <td class="field-value">{{{{TENANT_ADDRESS_DE}}}}</td>
        </tr>
        <tr>
          <td class="field-label">E-Mail / Tel:</td>
          <td class="field-value">{{{{TENANT_EMAIL}}}} | {{{{TENANT_PHONE}}}}</td>
        </tr>
      </table>
    </div>
    <div class="col">
      <div class="party-header">And <strong>tenant:</strong></div>
      <table class="field-table">
        <tr>
          <td class="field-label">Name:</td>
          <td class="field-value"><strong>{{{{TENANT_NAME}}}}</strong></td>
        </tr>
        <tr>
          <td class="field-label">Address:</td>
          <td class="field-value">{{{{TENANT_ADDRESS_EN}}}}</td>
        </tr>
        <tr>
          <td class="field-label">Email / Tel:</td>
          <td class="field-value">{{{{TENANT_EMAIL}}}} | {{{{TENANT_PHONE}}}}</td>
        </tr>
      </table>
    </div>
  </div>

  <!-- Specific Premises Breakdown Table -->
  <div class="premises-box">
    <div class="premises-title">
      Betreffend folgende Mietsache / Concerning the following premises:
    </div>
    <div class="premises-grid">
      <div><strong>Straße & Hausnummer / Street:</strong> Lichtenrader Damm 80</div>
      <div><strong>Geschoss & Whg / Floor & Flat:</strong> {{{{FLOOR}}}}. OG / Whg. WE-11</div>
      <div><strong>Zimmer-ID & Größe / Room:</strong> {{{{ROOM_CODE}}}} (ca. {{{{ROOM_SIZE}}}} m²)</div>
      <div><strong>Postleitzahl & Land / Zip & Country:</strong> 12305 Berlin, Deutschland / Germany</div>
      <div style="grid-column: span 2; margin-top: 3px; font-size: 8.2pt; color: #475569;">
        <strong>Gemeinschaftliche Flächen / Communal Amenities:</strong> Flur, voll ausgestattete Einbauküche, renoviertes Badezimmer/WC sowie Gepäckabteil dürfen vom Mieter vollumfänglich mitbenutzt werden. / Any communal amenities such as hallway, kitchen, shower/WC, and luggage room may be used by the tenant.
      </div>
    </div>
  </div>

  <!-- Preamble -->
  <div class="two-col">
    <div class="col">
      <div class="sec-title">Präambel</div>
      <div class="p-block">
        Der Vermieter vermietet möblierten Wohnraum zum vorübergehenden Gebrauch im Rahmen eines gemeinschaftlichen Co-Living-Konzepts. Der Mieter mietet das Zimmer zu befristeten Ausbildungs-, Studien- oder Arbeitszwecken in Berlin an.
      </div>
    </div>
    <div class="col">
      <div class="sec-title">Preamble</div>
      <div class="p-block">
        The landlord rents furnished living space for temporary use within a shared co-living arrangement. The tenant rents the room for a temporary educational, academic, or employment stay in Berlin.
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo_tag} &bull; Homelyo Rental Agreement</div>
    <div>Seite / Page 1</div>
  </div>
</div>

<!-- PAGE 2: Mietdauer, Miete, Internet & Kaution -->
<div class="page" id="page-2">
  <div class="page-header">
    <div class="header-left">{logo_img_tag}</div>
    <div class="header-right">Mietvertrag &bull; {{{{TENANT_NAME}}}} ({{{{ROOM_CODE}}}})</div>
  </div>

  <!-- § 1 Mietdauer -->
  <div class="two-col">
    <div class="col">
      <div class="sec-title">§ 1 Mietdauer & Befristung (§ 549 Abs. 2 BGB)</div>
      <div class="p-block">
        (1) Das Mietverhältnis beginnt am <strong>{{{{START_DATE_DE}}}}</strong> und endet am <strong>{{{{END_DATE_DE}}}}</strong>. Das Mietverhältnis ist fest befristet und verlängert sich nicht stillschweigend (§ 545 BGB ist abbedungen).
      </div>
      <div class="p-block">
        (2) <strong>Zweck der Befristung:</strong> Die Vermietung erfolgt zum vorübergehenden Gebrauch gem. § 549 Abs. 2 Nr. 1 BGB (Aufenthalt in Berlin zu: <em>{{{{REASON_DE}}}}</em>).
      </div>
      <div class="p-block">
        (3) <strong>Vorzeitige Entlassung / Nachmieter:</strong> Eine ordentliche Kündigung während der festen Vertragslaufzeit ist grundsätzlich ausgeschlossen. Der Mieter kann jedoch vorzeitig mit einer Frist von 3 Monaten zum Monatsende aus dem Vertrag entlassen werden, wenn er dem Vermieter einen geeigneten, solventen und für die WG zumutbaren Nachmieter stellt und dieser einen Mietvertrag abschließt.
      </div>
    </div>
    <div class="col">
      <div class="sec-title">§ 1 Rental Term & Temporary Purpose</div>
      <div class="p-block">
        (1) The tenancy commences on <strong>{{{{START_DATE_EN}}}}</strong> and terminates on <strong>{{{{END_DATE_EN}}}}</strong>. The lease is for a fixed term and tacit renewal is excluded (Section 545 BGB does not apply).
      </div>
      <div class="p-block">
        (2) <strong>Temporary Purpose:</strong> Concluded strictly for temporary living pursuant to Section 549 (2) No. 1 German Civil Code (BGB) for: <em>{{{{REASON_EN}}}}</em>.
      </div>
      <div class="p-block">
        (3) <strong>Early Release / Replacement Tenant:</strong> Ordinary early termination is excluded during the fixed lease. However, the tenant will be released early with 3 months' notice to month-end if they present an accepted, solvent, and compatible replacement tenant who enters into a formal binding lease.
      </div>
    </div>
  </div>

  <!-- § 2 Miete & Nebenkosten -->
  <div class="two-col">
    <div class="col">
      <div class="sec-title">§ 2 Miete, Nebenkosten & ARD</div>
      <div class="p-block">
        (1) Die monatliche Warmmiete beträgt pauschal <strong>{{{{RENT_AMOUNT}}}} €</strong> (in Worten: <strong>{{{{RENT_WORDS_DE}}}} EURO</strong>).
      </div>
      <div class="p-block">
        (2) In der Pauschalmiete enthalten sind alle kalten und warmen Betriebskosten (Heizung, Warm-/Kaltwasser, Strom, Müllabfuhr, Gebäudeversicherung) sowie Vodafone 1000 Mbit/s Highspeed WLAN. Eine jährliche Nebenkostenabrechnung findet nicht statt (Fair-Use-Regelung: Überschreitung &gt; 25% des Durchschnitts wird umgelegt).
      </div>
      <div class="p-block">
        (3) <span class="highlight-badge">WICHTIG &ndash; ARD / Rundfunkbeitrag:</span> <strong>Der Rundfunkbeitrag (GEZ) ist NICHT in der Miete enthalten.</strong> Die Anmeldung und Entrichtung des Rundfunkbeitrags muss von den Bewohnern der Wohngemeinschaft eigenständig vorgenommen, koordiniert und gezahlt werden.
      </div>
      <div class="p-block">
        (4) <strong>Einmalige Endreinigungspauschale:</strong> Für die professionelle Endreinigung des Zimmers und der anteiligen Gemeinschaftsräume bei Auszug wird eine einmalige Pauschale von <strong>70,00 €</strong> fällig.
      </div>
    </div>
    <div class="col">
      <div class="sec-title">§ 2 Rent, Utilities & Public License Fee</div>
      <div class="p-block">
        (1) The monthly all-inclusive warm rent is fixed at <strong>€{{{{RENT_AMOUNT}}}}</strong> (in words: <strong>{{{{RENT_WORDS_EN}}}} EURO</strong>).
      </div>
      <div class="p-block">
        (2) This flat rate covers all operational utilities (heating, hot/cold water, electricity, trash disposal, building insurance) and 1000 Mbps Wi-Fi 6. No annual reconciliation is performed (Fair use rule: usage &gt; 25% above property average may be charged).
      </div>
      <div class="p-block">
        (3) <span class="highlight-badge">IMPORTANT &ndash; ARD / TV-Radio License Fee:</span> <strong>The public broadcasting contribution (ARD / ZDF / GEZ) is NOT included in the rent.</strong> Registration and payment of the broadcasting license fee must be organized, coordinated, and paid directly by the tenants among themselves.
      </div>
      <div class="p-block">
        (4) <strong>One-Time Final Cleaning Fee:</strong> A one-time final cleaning fee of <strong>€70.00</strong> is payable for professional end-of-stay room and shared area sanitization upon departure.
      </div>
    </div>
  </div>

  <!-- § 3 Kaution & Zahlungsbedingungen -->
  <div class="two-col">
    <div class="col">
      <div class="sec-title">§ 3 Kaution & Zahlungsbedingungen</div>
      <div class="p-block">
        (1) <strong>Kaution:</strong> Der Mieter leistet eine Mietsicherheit in Höhe von <strong>{{{{DEPOSIT_AMOUNT}}}} €</strong> (in Worten: <strong>{{{{DEPOSIT_WORDS_DE}}}} EURO</strong>).
      </div>
      <div class="p-block">
        (2) <strong>Zahlung vor Einzug:</strong> Die Kaution sowie die erste Monatsmiete müssen <strong>vor dem Einzugsdatum ({{{{START_DATE_DE}}}})</strong> vollständig auf dem Vermieterkonto eingegangen sein. Ohne fristgerechten Geldeingang erfolgt keine Schlüsselübergabe und der Vertrag gilt als vom Mieter storniert.
      </div>
      <div class="p-block">
        (3) <strong>Mietfälligkeit:</strong> Die laufende Miete ist monatlich im Voraus <strong>bis spätestens zum 3. Werktag des Monats (vor dem 4.)</strong> kostenfrei auf das Vermieterkonto zu überweisen. Feiertage und Wochenenden sind bei der Überweisung im Voraus einzuplanen. Für jede verspätete Mietzahlung wird eine Mahngebühr von <strong>5,00 €</strong> fällig.
      </div>
    </div>
    <div class="col">
      <div class="sec-title">§ 3 Security Deposit & Payment Terms</div>
      <div class="p-block">
        (1) <strong>Security Deposit:</strong> The tenant pays a security deposit of <strong>€{{{{DEPOSIT_AMOUNT}}}}</strong> (in words: <strong>{{{{DEPOSIT_WORDS_EN}}}} EURO</strong>).
      </div>
      <div class="p-block">
        (2) <strong>Payment Prior to Move-In:</strong> Both the deposit and the first month's rent must be <strong>fully credited to the landlord's account before the move-in date ({{{{START_DATE_EN}}}})</strong>. Key handover will be denied without prior funds receipt, and the booking shall be deemed cancelled.
      </div>
      <div class="p-block">
        (3) <strong>Due Date:</strong> Ongoing rent is payable in advance <strong>on or before the 3rd business day of each month (before the 4th)</strong> via bank transfer. Bank holidays and weekends must be accounted for in advance. An overdue reminder fee of <strong>€5.00</strong> applies to each late payment notice.
      </div>
    </div>
  </div>

  <!-- Bank Details Box -->
  <div class="bank-box">
    <div class="bank-row">
      <span class="bank-label">Bankverbindung / Bank Account:</span>
      <span class="bank-val">N26 Bank AG, Berlin</span>
    </div>
    <div class="bank-row">
      <span class="bank-label">Kontoinhaber / Account Holder:</span>
      <span class="bank-val">{{{{RENT_ACCOUNT_HOLDER}}}}</span>
    </div>
    <div class="bank-row">
      <span class="bank-label">IBAN:</span>
      <span class="bank-val">{{{{RENT_IBAN}}}}</span>
    </div>
    <div class="bank-row">
      <span class="bank-label">BIC / SWIFT:</span>
      <span class="bank-val">{{{{RENT_BIC}}}}</span>
    </div>
    <div class="bank-row">
      <span class="bank-label">Verwendungszweck / Reference:</span>
      <span class="bank-val">{{{{TENANT_NAME}}}} - {{{{ROOM_CODE}}}} - Miete / Kaution</span>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo_tag} &bull; Homelyo Rental Agreement</div>
    <div>Seite / Page 2</div>
  </div>
</div>

<!-- PAGE 3: Internetverantwortung, Kleinreparaturen & Schlüsselgebühren -->
<div class="page" id="page-3">
  <div class="page-header">
    <div class="header-left">{logo_img_tag}</div>
    <div class="header-right">Mietvertrag &bull; {{{{TENANT_NAME}}}} ({{{{ROOM_CODE}}}})</div>
  </div>

  <!-- § 4 Internetverantwortung & Haftungsfreistellung -->
  <div class="two-col">
    <div class="col">
      <div class="sec-title">§ 4 Internetnutzung & Haftungsfreistellung</div>
      <div class="p-block">
        (1) Der Vermieter stellt dem Mieter einen WLAN-Zugang zur Verfügung. Der Mieter trägt die <strong>alleinige rechtliche Verantwortung (zivil- und strafrechtlich)</strong> für alle Aktivitäten, Downloads und Abrufe über diesen Anschluss.
      </div>
      <div class="p-block">
        (2) <strong>Verbot von Urheberrechtsverletzungen:</strong> Das Herunterladen, Anbieten oder Teilen von urheberrechtlich geschützten Inhalten über Peer-to-Peer-/Torrent-Netzwerke, illegale Streaming-Plattformen oder der Zugriff auf verbotene Webinhalte ist strengstens verboten.
      </div>
      <div class="p-block">
        (3) <strong>Freistellung & Sperrung:</strong> Der Mieter stellt den Vermieter von jeglichen Schadensersatzansprüchen, Abmahnungen und Rechtsanwaltskosten Dritter infolge rechtswidriger Internetnutzung vollumfänglich frei. Bei Verdacht auf illegale Nutzung ist der Vermieter berechtigt, den Internetzugang des Mieters fristlos zu sperren und die Mieterdaten an die ermittelnden Behörden zu übergeben.
      </div>
    </div>
    <div class="col">
      <div class="sec-title">§ 4 Internet Policy & Full Indemnification</div>
      <div class="p-block">
        (1) The landlord provides wireless internet (WLAN). The tenant carries <strong>sole and exclusive legal liability (both civil and criminal)</strong> for all data, downloads, and activities executed through the provided network.
      </div>
      <div class="p-block">
        (2) <strong>Copyright Infringement Prohibition:</strong> Downloading, uploading, or sharing copyrighted material via P2P/BitTorrent networks, unauthorized streaming portals, or visiting illicit sites is strictly forbidden.
      </div>
      <div class="p-block">
        (3) <strong>Indemnification & Access Revocation:</strong> The tenant fully indemnifies the landlord against any third-party claims, statutory warnings, and legal defense costs. In case of illegal use, the landlord reserves the right to immediately terminate the tenant's internet access and provide full identity and log records to prosecuting authorities.
      </div>
    </div>
  </div>

  <!-- § 5 Kleinreparaturen & Schlüsselverlust -->
  <div class="two-col">
    <div class="col">
      <div class="sec-title">§ 5 Kleinreparaturen & Schlüsselgebühren</div>
      <div class="p-block">
        (1) <strong>Kleinreparaturen (Bagatellschäden):</strong> Der Mieter trägt die Kosten für kleine Reparaturen an Gegenständen, die seinem häufigen und unmittelbaren Zugriff unterliegen (z.B. Wasserhähne, Lichtschalter, Steckdosen, Tür- und Fenstergriffe, WC-Spülung), sofern die Kosten im Einzelfall <strong>100,00 €</strong> nicht übersteigen. Die jährliche Gesamtbelastung ist auf 8% der Jahresmiete, maximal 350,00 € pro Kalenderjahr begrenzt.
      </div>
      <div class="p-block">
        (2) <strong>Schlüsselverlust-Gebührensatz:</strong> Bei Verlust oder Beschädigung von Schlüsseln hat der Mieter die tatsächlichen Ersatzkosten zu tragen:
      </div>
      <div class="p-block" style="font-size: 8.0pt; margin-left: 8px;">
        &bull; <strong>Haustürschlüssel (Gebäudeeingang): 680,00 €</strong><br>
        &bull; <strong>Wohnungstürschlüssel: 120,00 €</strong><br>
        &bull; <strong>Zimmertürschlüssel: 90,00 €</strong><br>
        &bull; <strong>Briefkastenschlüssel: 70,00 €</strong><br>
        Ist aus Sicherheitsgründen der Austausch von Schlössern oder der Schließanlage erforderlich, haftet der Mieter vollumfänglich für die Gesamtkosten.
      </div>
    </div>
    <div class="col">
      <div class="sec-title">§ 5 Minor Repairs & Key Replacement Schedule</div>
      <div class="p-block">
        (1) <strong>Minor Repairs Clause:</strong> The tenant pays for minor repairs to installations subject to frequent and direct use (e.g. water faucets, light switches, door/window handles, toilet flushes) up to <strong>€100.00</strong> per single repair. Annual liability is capped at 8% of the annual rent, up to a maximum of €350.00 per calendar year.
      </div>
      <div class="p-block">
        (2) <strong>Key Replacement Fee Schedule:</strong> In the event of lost or damaged keys, the tenant compensates the actual replacement and administrative costs:
      </div>
      <div class="p-block" style="font-size: 8.0pt; margin-left: 8px;">
        &bull; <strong>Building entrance door key: €680.00</strong><br>
        &bull; <strong>Apartment entrance key: €120.00</strong><br>
        &bull; <strong>Private room key: €90.00</strong><br>
        &bull; <strong>Mailbox key: €70.00</strong><br>
        If security requires replacing cylinder locks or the central master lock system, the tenant is liable for the full replacement expense.
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo_tag} &bull; Homelyo Rental Agreement</div>
    <div>Seite / Page 3</div>
  </div>
</div>

<!-- ==================== PART 2: ALLGEMEINE GESCHÄFTSBEDINGUNGEN (AGB) ==================== -->

<!-- PAGE 4: AGB § 1 Zusammenleben in der WG -->
<div class="page" id="page-4">
  <div class="page-header">
    <div class="header-left">{logo_img_tag}</div>
    <div class="header-right">Allgemeine Geschäftsbedingungen (AGB) &bull; Homelyo</div>
  </div>

  <div class="agb-header">
    <h3>Allgemeine Geschäftsbedingungen (AGB) der Wohngemeinschaft</h3>
    <h4>General Terms and Conditions (GTC) for Flat Share Cohabitation</h4>
  </div>

  <!-- AGB Preamble -->
  <div class="two-col">
    <div class="col">
      <div class="p-block">
        Der Mieter erkennt die Allgemeinen Geschäftsbedingungen (AGB) als für ihn verbindlich an. Ein Verstoß gegen die AGB ist ein vertragswidriger Gebrauch des Mietgegenstandes. Bei schwerwiegenden Fällen kann der Vermieter das Mietverhältnis ohne Einhaltung der Kündigungsfrist kündigen. Für alle Schäden, die dem Vermieter durch Verletzung oder Nichtbeachtung der AGB sowie durch Nichterfüllung der Meldepflichten entstehen, ist der Mieter ersatzpflichtig. Die AGB regeln das Zusammenleben aller Mitbewohner der Wohnung und des Hauses. Sie enthalten Rechte und Pflichten und gelten für alle Bewohner der Wohngemeinschaft gleichermaßen.
      </div>
    </div>
    <div class="col">
      <div class="p-block">
        The tenant accepts these General Terms and Conditions (GTC) as obligatory. A violation of the terms is considered a breach of contract and misuse of the rented property. In severe cases the lessor is entitled to terminate the contract without notice. The tenant is obliged to financially compensate for any and all damage to the property caused by neglect or violation of the agreed terms as well as not upholding reporting requirements. The GTC regulates the cohabitation of all residents of the apartment and building equally.
      </div>
    </div>
  </div>

  <!-- § 1 Zusammenleben -->
  <div class="two-col">
    <div class="col">
      <div class="sec-title">§ 1 Zusammenleben in der WG / mit den Nachbarn</div>
      <div class="p-block">
        <strong>1. Rücksicht:</strong> Die WG hat das Zusammenleben untereinander mit Rücksicht auf die Interessen und zur Zufriedenheit aller Beteiligten fair und eigenständig zu regeln.
      </div>
      <div class="p-block">
        <strong>2. Lärm & Ruhezeiten:</strong> Jeder Mieter ist dafür verantwortlich, dass vermeidbarer Lärm unterbleibt. Besondere Rücksichtnahme gilt in der Mittagsruhe von <strong>13:00 bis 15:00 Uhr</strong> sowie zur Nachtruhe zwischen <strong>22:00 und 08:00 Uhr</strong>. An Sonn- und Feiertagen ist Lärm komplett zu unterlassen. TV, Soundanlagen etc. sind auf Zimmerlautstärke zu halten. Das Musizieren ist während der Ruhezeiten untersagt und außerhalb auf maximal 2 Stunden täglich beschränkt.
      </div>
      <div class="p-block">
        <strong>3. Partys:</strong> Störende Partys sind komplett untersagt. Bei geselligem Beisammensein ist streng auf die Lautstärke zu achten, die spätestens um 22:00 Uhr auf Zimmerlautstärke zu senken ist.
      </div>
      <div class="p-block">
        <strong>4. Reinigung:</strong> Die Mieter reinigen die Mieträume einschließlich der Gemeinschaftsräume (Küche, Bad, WC, Flur, Balkon) eigenverantwortlich und halten diese frei von Schmutz und Unrat. Erfolgt trotz Fristsetzung keine Reinigung, kann der Vermieter eine professionelle Reinigung auf Kosten aller Mieter (zu gleichen Teilen) beauftragen. Wiederholte Nichtbeteiligung berechtigt nach Abmahnung zur fristlosen Kündigung.
      </div>
      <div class="p-block">
        <strong>5. Rauch- & Drogenverbot:</strong> Das Rauchen ist in der gesamten Wohnung, im Treppenhaus und auf den Balkonen strengstens verboten. Besitz und Konsum illegaler Drogen auf dem gesamten Grundstück sind verboten, werden ausnahmslos polizeilich zur Anzeige gebracht und berechtigen zur fristlosen Kündigung.
      </div>
      <div class="p-block">
        <strong>6. Haustiere:</strong> Die Haltung von Haustieren ist grundsätzlich untersagt.
      </div>
      <div class="p-block">
        <strong>7. Beherbergung Dritter:</strong> Eine Gebrauchsüberlassung an Dritte oder die dauerhafte Aufnahme von Lebensgefährten ist unzulässig. Der Besuch von Familie/Freunden ist auf normale Besuche beschränkt. Ein Aufenthalt von Besuchern an mehr als fünf (5) Tagen im Monat gilt als vertragswidrige Nutzung und berechtigt zur fristlosen Kündigung. Besuch ist nur bei persönlicher Anwesenheit des Mieters gestattet.
      </div>
    </div>
    <div class="col">
      <div class="sec-title">§ 1 Cohabitation in Flat Share / Neighbours</div>
      <div class="p-block">
        <strong>1. Regard & Consideration:</strong> Flatmates must organize cohabitation independently with mutual respect and fairness toward each other.
      </div>
      <div class="p-block">
        <strong>2. Noise & Quiet Hours:</strong> Tenants must prevent unnecessary noise. Mandatory quiet hours are observed during midday from <strong>1:00pm to 3:00pm</strong> and at night between <strong>10:00pm and 8:00am</strong>. Excessive noise is strictly forbidden on Sundays and public holidays. Media devices must operate at room volume. Instrument practice is prohibited during quiet hours and limited to 2 hours daily otherwise.
      </div>
      <div class="p-block">
        <strong>3. Parties:</strong> Disruptive parties are strictly prohibited. Social gatherings must be kept quiet and reduced to room volume by 10:00pm.
      </div>
      <div class="p-block">
        <strong>4. Cleaning & Maintenance:</strong> Tenants clean their private rooms and all communal areas (kitchen, bath, WC, hall, balcony) under their own responsibility. If cleaning is neglected despite a warning deadline, the landlord will hire professional cleaners at the equal expense of all flatmates. Repeated refusal leads to lease termination.
      </div>
      <div class="p-block">
        <strong>5. Smoking & Substance Ban:</strong> Smoking is strictly forbidden throughout the entire apartment and building. Possession or consumption of illegal narcotics is prohibited, reported to police immediately, and results in immediate lease cancellation.
      </div>
      <div class="p-block">
        <strong>6. Pets:</strong> Keeping animals or pets of any kind is prohibited.
      </div>
      <div class="p-block">
        <strong>7. Overnight Guests:</strong> Subletting or permanently lodging partners/spouses is prohibited. Visitors exceeding five (5) days in a single month constitute contract breach and permit termination without notice. Guests are only permitted while the tenant is personally present.
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo_tag} &bull; Homelyo GTC / AGB</div>
    <div>Seite / Page 4</div>
  </div>
</div>

<!-- PAGE 5: AGB § 2 Regelungen zur Wohnung -->
<div class="page" id="page-5">
  <div class="page-header">
    <div class="header-left">{logo_img_tag}</div>
    <div class="header-right">Allgemeine Geschäftsbedingungen (AGB) &bull; Homelyo</div>
  </div>

  <div class="two-col">
    <div class="col">
      <div class="sec-title">§ 2 Regelungen zur Wohnung</div>
      <div class="p-block">
        <strong>1. Nutzung:</strong> Die Anmietung erfolgt ausschließlich zu Wohnzwecken. Gewerbliche Nutzung ist unzulässig.
      </div>
      <div class="p-block">
        <strong>2. Lüften und Heizen:</strong> Die Wohnung ist auch in der kalten Jahreszeit mindestens zweimal täglich durch kurzes, vollständiges Öffnen der Fenster (Stoßlüften) zu lüften. Beim Lüften und bei längerer Abwesenheit sind die Heizkörperthermostate herunterzudrehen.
      </div>
      <div class="p-block">
        <strong>3. Internet & Rundfunk:</strong> Eigene Telefonanschlüsse dürfen nicht installiert werden. W-LAN wird zur Verfügung gestellt. Der Mieter trägt die volle Verantwortung für sein Handeln im Netz; illegale Downloads führen zu strafrechtlicher Verfolgung mit Weitergabe der Mieterdaten. Der Rundfunkbeitrag (GEZ) ist von der WG selbstständig anzumelden und aufzuteilen.
      </div>
      <div class="p-block">
        <strong>4. Bilder und Plakate:</strong> Das Anbringen von Bildern, Regalen etc. mit Nägeln, Schrauben oder Bohrungen (Dübeln) ist untersagt. Eigene Schilder oder Aufschriften an Haus- und Wohnungstüren sind verboten.
      </div>
      <div class="p-block">
        <strong>5. Obhut & Sorgfalt:</strong> Fette, Speisereste und Hygieneartikel dürfen nicht in die Abflüsse (Küche, Bad, WC) gelangen. Böden sind pfleglich zu behandeln und vor Kratzern zu schützen.
      </div>
      <div class="p-block">
        <strong>6. Müllentsorgung:</strong> Der Müll ist konsequent nach den Vorgaben der Berliner Stadtreinigung (BSR) zu trennen (Plastik/Wertstoffe, Papier, Bio, Restmüll). Sperrmüll und Sondermüll gehören nicht in die Hausmülltonnen.
      </div>
      <div class="p-block">
        <strong>7. Bauliche Veränderungen:</strong> Bauliche Veränderungen und das Anbringen von Satellitenempfangsanlagen sind strengstens untersagt.
      </div>
      <div class="p-block">
        <strong>8. Sicherheit:</strong> Haustüren, Hof- und Kellertüren sind zwischen 22:00 und 06:00 Uhr geschlossen zu halten. Fluchtwege (Flure, Treppen) sind stets freizuhalten. Das Grillen auf den Balkonen ist verboten.
      </div>
      <div class="p-block">
        <strong>9. Betretungsrecht:</strong> Der Vermieter bzw. Beauftragte dürfen die Mieträume nach Vorankündigung werktags zwischen 09:00-12:00 Uhr und 15:00-20:00 Uhr zur Überprüfung, für Besichtigungen mit Mietinteressenten (in den letzten 4 Wochen mind. 5 Termine) sowie bei Gefahr im Verzug jederzeit betreten. Zur Überprüfung des Hygiene- und Reinigungszustands der Gemeinschaftsflächen darf ein externer Dienstleister beauftragt werden.
      </div>
    </div>
    <div class="col">
      <div class="sec-title">§ 2 Apartment Regulations</div>
      <div class="p-block">
        <strong>1. Residential Use:</strong> Strictly for private residential purposes. Commercial activity is prohibited.
      </div>
      <div class="p-block">
        <strong>2. Ventilation & Heating:</strong> Air out rooms at least twice daily by opening windows fully for a brief period. Turn down radiators when airing or when absent.
      </div>
      <div class="p-block">
        <strong>3. Internet & Broadcasting:</strong> Private phone lines may not be installed. WLAN is provided. The tenant carries full responsibility for lawful internet use. The public broadcasting license fee (ARD) must be registered and coordinated by tenants directly.
      </div>
      <div class="p-block">
        <strong>4. Walls & Posters:</strong> Drilling, screws, nails, or wall plugs for shelves/pictures are strictly forbidden. Attaching personal exterior signs or stickers is prohibited.
      </div>
      <div class="p-block">
        <strong>5. Care of Drains & Floors:</strong> No food waste, oils, grease, or solids down kitchen/bathroom drains. Parquet/laminate flooring must be protected from scratches.
      </div>
      <div class="p-block">
        <strong>6. Waste Separation:</strong> Waste must be systematically sorted (packaging/plastics, paper, organic, residual waste) according to municipal regulations.
      </div>
      <div class="p-block">
        <strong>7. Modifications:</strong> Structural modifications or installing exterior satellite dishes are prohibited.
      </div>
      <div class="p-block">
        <strong>8. Security:</strong> Building and yard doors must remain locked between 10:00pm and 06:00am. Hallways and escape paths must remain unobstructed. Grilling on balconies is prohibited.
      </div>
      <div class="p-block">
        <strong>9. Inspection & Entry:</strong> The landlord or designated contractors may enter on workdays between 09:00am-12:00pm and 3:00pm-8:00pm upon notice for evaluation, maintenance, and prospective tenant viewings (at least 5 viewings during the last 4 weeks of stay), and immediately in case of emergency. Cleaning condition inspections of common areas may be performed.
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo_tag} &bull; Homelyo GTC / AGB</div>
    <div>Seite / Page 5</div>
  </div>
</div>

<!-- PAGE 6: AGB § 3 Zahlungen, § 4 Übergabe & § 5 Schäden -->
<div class="page" id="page-6">
  <div class="page-header">
    <div class="header-left">{logo_img_tag}</div>
    <div class="header-right">Allgemeine Geschäftsbedingungen (AGB) &bull; Homelyo</div>
  </div>

  <div class="two-col">
    <!-- German Column -->
    <div class="col">
      <div class="sec-title">§ 3 Zahlungsvereinbarungen</div>
      <div class="p-block">
        <strong>1. Rechtzeitigkeit:</strong> Maßgeblich für die Rechtzeitigkeit von Miet- und Kautionszahlungen ist das Datum des Zahlungseingangs auf dem Vermieterkonto, nicht das Absendedatum.
      </div>
      <div class="p-block">
        <strong>2. Kaution:</strong> Dient der Sicherung aller Ansprüche (Miete, Schäden, Schlüsselverlust). Sie ist vor Einzug per Überweisung zu zahlen. Eine Aufrechnung gegen laufende Miete vor Mietende ist ausgeschlossen.
      </div>
      <div class="p-block">
        <strong>3. Rückerstattung:</strong> Nach Beendigung des Mietverhältnisses und ordnungsgemäßer Rückgabe der Mietsache wird die Kaution abzüglich berechtigter Forderungen innerhalb von maximal 30 Tagen per Banküberweisung erstattet.
      </div>

      <div class="sec-title">§ 4 Übergabe und Rückgabe</div>
      <div class="p-block">
        <strong>1. Zeitpunkt:</strong> Übergabe erfolgt erst nach vollständigem Eingang von Kaution und erster Miete.
      </div>
      <div class="p-block">
        <strong>2. Übergabeprotokoll:</strong> Vor Einzug wird ein Protokoll angefertigt. Das Zimmer muss bei Auszug im selben Zustand besenrein und frei von persönlichen Gegenständen zurückgegeben werden.
      </div>
      <div class="p-block">
        <strong>3. Schlüsselverlust:</strong> Bei Schlüsselverlust haftet der Mieter für: <strong>Haustür 680 €, Wohnung 120 €, Zimmer 90 €, Briefkasten 70 €</strong> sowie ggf. erforderliche Schlossänderungen der Schließanlage.
      </div>

      <div class="sec-title">§ 5 Schäden an der Wohnung</div>
      <div class="p-block">
        <strong>1. Anzeigepflicht:</strong> Mängel sind dem Vermieter unverzüglich anzuzeigen. Dem Vermieter ist eine Frist von 4 Wochen zur Abhilfe einzuräumen.
      </div>
      <div class="p-block">
        <strong>2. Haftung & Beweislast:</strong> Der Mieter haftet für Schäden durch Verletzung seiner Sorgfaltspflichten sowie das Verschulden seiner Besucher. Im Zimmer und in den Gemeinschaftsräumen trägt der Mieter die Beweislast für fehlendes Verschulden.
      </div>
      <div class="p-block">
        <strong>3. Schönheitsreparaturen & Kleinreparaturen:</strong> Mieter schuldet keine routinemäßigen Schönheitsreparaturen. Kleinreparaturen bis 100 € je Maßnahme (max. 8% Jahresmiete / 350 € jährlich) trägt der Mieter.
      </div>
    </div>

    <!-- English Column -->
    <div class="col">
      <div class="sec-title">§ 3 Payment Agreements</div>
      <div class="p-block">
        <strong>1. Punctuality:</strong> Payment punctuality is determined exclusively by the date funds are credited to the landlord's account, not submission date.
      </div>
      <div class="p-block">
        <strong>2. Deposit:</strong> Secures all claims (rent, damages, lost keys). Must be paid prior to handover. Setting off the deposit against ongoing rent is prohibited.
      </div>
      <div class="p-block">
        <strong>3. Refund:</strong> Returned within 30 days after lease expiry, property return, and signed checkout protocol, less proven damages or claims.
      </div>

      <div class="sec-title">§ 4 Handover and Return</div>
      <div class="p-block">
        <strong>1. Handover Timing:</strong> Carried out only after verified receipt of security deposit and first rent.
      </div>
      <div class="p-block">
        <strong>2. Protocol:</strong> An inventory handover protocol is signed at check-in. The room must be returned swept clean and cleared of all belongings at departure.
      </div>
      <div class="p-block">
        <strong>3. Keys:</strong> Tenant compensates lost keys at: <strong>Building entrance €680, Apartment €120, Room €90, Mailbox €70</strong>, plus master cylinder lock changes.
      </div>

      <div class="sec-title">§ 5 Damages to Property</div>
      <div class="p-block">
        <strong>1. Notification:</strong> Defects must be reported immediately; landlord has 4 weeks to arrange remediation.
      </div>
      <div class="p-block">
        <strong>2. Liability:</strong> The tenant is liable for neglect and visitor damage. The burden of proof rests with the tenant for damage arising within their rooms.
      </div>
      <div class="p-block">
        <strong>3. Maintenance & Minor Repairs:</strong> No general cosmetic renovation required. Minor repairs under €100 per instance (up to 8% annual rent / €350 cap) are tenant's responsibility.
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo_tag} &bull; Homelyo GTC / AGB</div>
    <div>Seite / Page 6</div>
  </div>
</div>

<!-- PAGE 7: AGB § 6 Kündigung, § 7 Schriftform, § 8 Salvatorische Klausel & Signatures -->
<div class="page" id="page-7">
  <div class="page-header">
    <div class="header-left">{logo_img_tag}</div>
    <div class="header-right">Allgemeine Geschäftsbedingungen (AGB) &amp; Unterschriften</div>
  </div>

  <div class="two-col">
    <!-- German Column -->
    <div class="col">
      <div class="sec-title">§ 6 Kündigungsbedingungen</div>
      <div class="p-block">
        <strong>1. Nachmieterstellung:</strong> Schlägt der Mieter mit einer Frist von 3 Monaten zum Monatsende einen geeigneten Nachmieter vor und wird dieser vom Vermieter akzeptiert, wird der Mieter ab Übergabe an den Nachmieter von seiner Zahlungspflicht befreit. Andernfalls haftet der Mieter bis zum regulären Ablauf der vereinbarten Mietzeit.
      </div>
      <div class="p-block">
        <strong>2. Fristlose Kündigung:</strong> Der Vermieter behält sich die fristlose Kündigung vor bei: Verstoß gegen AGB trotz Abmahnung, vom Mieter verursachten Schäden &gt; 50 €, unbefugter Gebrauchsüberlassung an Dritte, beharrlicher Reinigungsverweigerung, Erhalt von 2 Abmahnungen oder Zahlungsverzug von zwei Monatsmieten (§ 543 Abs. 2 BGB).
      </div>
      <div class="p-block">
        <strong>3. Ausschluss stillschweigender Verlängerung:</strong> § 545 BGB wird ausdrücklich ausgeschlossen.
      </div>

      <div class="sec-title">§ 7 Schriftform &amp; § 8 Salvatorische Klausel</div>
      <div class="p-block">
        Mündliche Nebenabreden bestehen nicht. Änderungen bedürfen der Schriftform. Sollten einzelne Bestimmungen unwirksam sein, berührt dies die Wirksamkeit der übrigen AGB nicht. Anstelle der unwirksamen Regelung gilt diejenige Bestimmung, die dem wirtschaftlichen Zweck am nächsten kommt.
      </div>
    </div>

    <!-- English Column -->
    <div class="col">
      <div class="sec-title">§ 6 Termination Conditions</div>
      <div class="p-block">
        <strong>1. Replacement Tenant:</strong> If the tenant presents an accepted solvent replacement tenant with 3 months' notice to month-end who executes a contract, the tenant is released from rent obligations from the replacement move-in date. Otherwise, the tenant remains fully liable for the rent until the scheduled lease end.
      </div>
      <div class="p-block">
        <strong>2. Termination Without Notice:</strong> Landlord may terminate immediately for: persistent breach of GTC despite warning, damages &gt; €50 unpaid, unauthorized third-party lodging, failure to maintain cleanliness after warning, receiving 2 written warnings, or 2 months rent default (§ 543 (2) BGB).
      </div>
      <div class="p-block">
        <strong>3. Exclusion of Tacit Renewal:</strong> Section 545 BGB is strictly excluded.
      </div>

      <div class="sec-title">§ 7 Written Form &amp; § 8 Severability Clause</div>
      <div class="p-block">
        No oral side agreements exist. Amendments must be in writing. If any provision is deemed invalid, the remaining terms retain full force. The invalid clause will be replaced with a provision closest to the intended economic purpose.
      </div>
    </div>
  </div>

  <!-- Signatures Section with NO timestamps -->
  <div class="sig-block">
    <div class="two-col">
      <!-- Landlord -->
      <div class="col">
        <div class="sig-line"></div>
        <div class="sig-title">{{{{LANDLORD_SIGN_NAME}}}}</div>
        <div class="sig-date">Vermieter / Landlord &bull; Datum / Date: ____________________</div>
      </div>

      <!-- Tenant -->
      <div class="col">
        <div class="sig-line"></div>
        <div class="sig-title">{{{{TENANT_SIGN_NAME}}}}</div>
        <div class="sig-date">Mieter / Tenant &bull; Datum / Date: ____________________</div>
      </div>
    </div>
  </div>

  <div class="page-footer">
    <div>{footer_logo_tag} &bull; Homelyo GTC &amp; Signatures</div>
    <div>Seite / Page 7</div>
  </div>
</div>

</body>
</html>
"""
    return html

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
    individual_de=None,
    individual_en=None,
    landlord_sign_date=None,
    tenant_sign_date=None
):
    spec = ROOM_SPECS.get(room_code, ROOM_SPECS["KE2"])
    rent = float(rent_amount) if rent_amount is not None else spec["rent"]
    deposit = float(deposit_amount) if deposit_amount is not None else 1000.00

    rent_str = f"{int(rent)}" if rent.is_integer() else f"{rent:.2f}".replace(".", ",")
    deposit_str = f"{int(deposit):,}".replace(",", ".") if deposit.is_integer() else f"{deposit:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    rent_words_de = num_to_words_de(rent)
    rent_words_en = num_to_words_en(rent)
    deposit_words_de = num_to_words_de(deposit)
    deposit_words_en = num_to_words_en(deposit)

    start_de, start_en = format_date_bilingual(start_date)
    end_de, end_en = format_date_bilingual(end_date)

    if not tenant_address_de:
        tenant_address_de = "Gemäß Personalausweis / Passdokument"
    if not tenant_address_en:
        tenant_address_en = "As per identity card / passport document"

    if not reason_de:
        reason_de = "Aufgrund eines Studien-, Praktikums- oder Erwerbsaufenthalts in Berlin"
    if not reason_en:
        reason_en = "Due to an academic, internship, or professional stay in Berlin"

    return {
        "LANDLORD_NAME": "Kylasam NA",
        "LANDLORD_SIGN_NAME": "Kylasam Nallagoundanpalayam Arruchamy",
        "LANDLORD_ADDRESS_DE": "Forckenbeckstraße 64, 14199 Berlin, Deutschland",
        "LANDLORD_ADDRESS_EN": "Forckenbeckstraße 64, 14199 Berlin, Germany",
        "PROPERTY_ADDRESS_DE": "Lichtenrader Damm 80, 12305 Berlin, Deutschland",
        "PROPERTY_ADDRESS_EN": "Lichtenrader Damm 80, 12305 Berlin, Germany",
        "FLOOR": spec.get("floor", "2"),
        "ROOM_COUNT": spec.get("room_num", "1"),
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
        "REASON_DE": reason_de,
        "REASON_EN": reason_en,
        "RENT_AMOUNT": rent_str,
        "RENT_WORDS_DE": rent_words_de,
        "RENT_WORDS_EN": rent_words_en,
        "DEPOSIT_AMOUNT": deposit_str,
        "DEPOSIT_WORDS_DE": deposit_words_de,
        "DEPOSIT_WORDS_EN": deposit_words_en,
        "FIRST_RENT_AMOUNT": rent_str,
        "CLEANING_FEE": "70",
        "RENT_ACCOUNT_HOLDER": "Kylasam Nallagoundanpalayam Arruchamy",
        "RENT_IBAN": "DE87100110012624542768",
        "RENT_BIC": "NTSBDEB1XXX"
    }

def render_contract_html(data):
    html = generate_html_template()
    for k, v in data.items():
        html = html.replace(f"{{{{{k}}}}}", str(v))
    return html

def render_contract_docx(data, output_path):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.55)
        section.right_margin = Inches(0.55)
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
    r_de.font.size = Pt(24)
    r_de.font.bold = True
    r_en = p_title.add_run("Residential Lease Agreement & House Rules (AGB)")
    r_en.font.size = Pt(13)
    r_en.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Parties Table
    t_parties = doc.add_table(rows=4, cols=2)
    t_parties.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r in t_parties.rows:
        r.cells[0].width = Inches(3.5)
        r.cells[1].width = Inches(3.5)

    t_parties.rows[0].cells[0].paragraphs[0].add_run("Vermieter / Landlord: ").font.bold = True
    t_parties.rows[0].cells[0].paragraphs[0].add_run(f"{data['LANDLORD_SIGN_NAME']}\n{data['LANDLORD_ADDRESS_DE']}")
    t_parties.rows[0].cells[1].paragraphs[0].add_run("Mieter / Tenant: ").font.bold = True
    t_parties.rows[0].cells[1].paragraphs[0].add_run(f"{data['TENANT_NAME']}\n{data['TENANT_ADDRESS_DE']}\n{data['TENANT_EMAIL']} | {data['TENANT_PHONE']}")

    t_parties.rows[1].cells[0].paragraphs[0].add_run("Mietsache / Premises:\n").font.bold = True
    t_parties.rows[1].cells[0].paragraphs[0].add_run(f"Zimmer {data['ROOM_CODE']} (ca. {data['ROOM_SIZE']} m²)\nLichtenrader Damm 80, 12305 Berlin\nGeschoss: {data['FLOOR']}. OG, Whg. WE-11")
    t_parties.rows[1].cells[1].paragraphs[0].add_run("Mietdauer & Zweck / Term:\n").font.bold = True
    t_parties.rows[1].cells[1].paragraphs[0].add_run(f"{data['START_DATE_DE']} bis {data['END_DATE_DE']}\nBefristet gem. § 549 Abs. 2 BGB:\n{data['REASON_DE']}")

    t_parties.rows[2].cells[0].paragraphs[0].add_run("Miete & Nebenkosten:\n").font.bold = True
    t_parties.rows[2].cells[0].paragraphs[0].add_run(f"Monatlich {data['RENT_AMOUNT']} € ({data['RENT_WORDS_DE']} EURO)\nInkl. Nebenkosten & 1000 Mbit/s WLAN\nEndreinigung: {data['CLEANING_FEE']} €\nARD-Rundfunkbeitrag: Von WG zu tragen")
    t_parties.rows[2].cells[1].paragraphs[0].add_run("Rent & Utilities:\n").font.bold = True
    t_parties.rows[2].cells[1].paragraphs[0].add_run(f"Monthly €{data['RENT_AMOUNT']} ({data['RENT_WORDS_EN']} EURO)\nIncl. utilities & 1000 Mbps Wi-Fi\nFinal cleaning: €{data['CLEANING_FEE']}\nARD public license fee: Paid by flatmates")

    t_parties.rows[3].cells[0].paragraphs[0].add_run("Kaution & Bankkonto:\n").font.bold = True
    t_parties.rows[3].cells[0].paragraphs[0].add_run(f"Kaution: {data['DEPOSIT_AMOUNT']} € (Zahlung vor Einzug)\nIBAN: {data['RENT_IBAN']}\nBIC: {data['RENT_BIC']}\nInhaber: {data['RENT_ACCOUNT_HOLDER']}")
    t_parties.rows[3].cells[1].paragraphs[0].add_run("Deposit & Bank Account:\n").font.bold = True
    t_parties.rows[3].cells[1].paragraphs[0].add_run(f"Deposit: €{data['DEPOSIT_AMOUNT']} (credited prior to move-in)\nIBAN: {data['RENT_IBAN']}\nBIC: {data['RENT_BIC']}\nHolder: {data['RENT_ACCOUNT_HOLDER']}")

    doc.add_page_break()

    p_agb_title = doc.add_paragraph()
    p_agb_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_agb_de = p_agb_title.add_run("Allgemeine Mietvertragsbestimmungen & Hausordnung (§ 1 – § 14)\n")
    r_agb_de.font.size = Pt(14)
    r_agb_de.font.bold = True
    r_agb_en = p_agb_title.add_run("General Lease Terms & House Rules (AGB § 1 – § 14)")
    r_agb_en.font.size = Pt(11)
    r_agb_en.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 14 Full Clauses Table (§ 1 to § 14)
    clauses_data = [
        (
            "§ 1 Mietdauer & Befristungszweck (§ 549 Abs. 2 / § 575 BGB)\n",
            f"(1) Mietzeitraum: Das Mietverhältnis beginnt am {data['START_DATE_DE']} und endet am {data['END_DATE_DE']}. Das Mietverhältnis ist befristet. Eine stillschweigende Verlängerung nach § 545 BGB ist ausdrücklich ausgeschlossen.\n"
            f"(2) Zweck der Befristung: Die Vermietung erfolgt zum vorübergehenden Gebrauch gem. § 549 Abs. 2 Nr. 1 BGB ({data['REASON_DE']}).\n"
            "(3) Mindestmietdauer: Die Mindestmietdauer beträgt sechs (6) Monate (bzw. die vereinbarte feste Vertragslaufzeit).",
            "§ 1 Lease Term & Temporary Purpose (§ 549 (2) / § 575 BGB)\n",
            f"(1) Rental Term: Tenancy commences on {data['START_DATE_EN']} and terminates on {data['END_DATE_EN']}. The lease is for a fixed term. Tacit renewal pursuant to Section 545 BGB is expressly excluded.\n"
            f"(2) Purpose of Limitation: Concluded strictly for temporary living pursuant to Section 549 (2) No. 1 BGB ({data['REASON_EN']}).\n"
            "(3) Minimum Term: The minimum rental term is six (6) months (or the agreed contract period)."
        ),
        (
            "§ 2 Miethöhe, Pauschale & Nebenkosten\n",
            f"(1) Monatliche Gesamtmiete: Die monatliche Warmmiete beträgt pauschal {data['RENT_AMOUNT']} € ({data['RENT_WORDS_DE']} EURO).\n"
            "(2) Inklusivleistungen: In der Pauschalmiete enthalten sind alle kalten und warmen Betriebskosten (Heizung, Warm-/Kaltwasser, Strom, Müllabfuhr, Gebäudeversicherung) sowie Vodafone 1000 Mbit/s Highspeed-WLAN in verbrauchsüblichem Maße (Fair-Use-Regelung).\n"
            f"(3) ARD-Rundfunkbeitrag & Endreinigung: Der Rundfunkbeitrag (GEZ) ist NICHT in der Miete enthalten und wird von der WG eigenständig getragen. Für die Endreinigung bei Auszug wird eine einmalige Pauschale von {data['CLEANING_FEE']} € fällig.",
            "§ 2 Rent, Utilities & Public License Fee\n",
            f"(1) Monthly Flat Rent: Monthly warm rent is an all-inclusive flat rate of €{data['RENT_AMOUNT']} ({data['RENT_WORDS_EN']} EURO).\n"
            "(2) Inclusions: Includes all operating costs (heating, hot/cold water, electricity, waste disposal, insurance) and Vodafone 1000 Mbps Wi-Fi subject to customary fair use.\n"
            f"(3) ARD License & Cleaning: Public broadcasting license (GEZ) is NOT included and must be settled by tenants. A one-time final checkout cleaning fee of €{data['CLEANING_FEE']} applies upon departure."
        ),
        (
            "§ 3 Kaution & Mietsicherheit (§ 551 BGB)\n",
            f"(1) Kaution: Der Mieter leistet eine Kaution von {data['DEPOSIT_AMOUNT']} € ({data['DEPOSIT_WORDS_DE']} EURO) zur Sicherung aller Ansprüche.\n"
            f"(2) Fälligkeit vor Einzug: Kaution und 1. Monatsmiete müssen vor dem Einzugstermin ({data['START_DATE_DE']}) vollständig auf dem Vermieterkonto eingegangen sein. Ohne Geldeingang erfolgt keine Schlüsselübergabe.\n"
            "(3) Rückerstattung: Treuhänderische Verwahrung gem. § 551 BGB. Erstattung innerhalb von 30 Tagen nach Auszug und unterzeichnetem Protokoll abzüglich berechtigter Forderungen. Aufrechnung gegen die letzte Miete ist ausgeschlossen.",
            "§ 3 Security Deposit & Escrow (§ 551 BGB)\n",
            f"(1) Deposit: Tenant renders a deposit of €{data['DEPOSIT_AMOUNT']} ({data['DEPOSIT_WORDS_EN']} EURO) to secure all landlord claims.\n"
            f"(2) Due Prior to Move-In: Deposit and 1st month's rent must be fully credited before move-in date ({data['START_DATE_EN']}). Handover of keys is denied without full receipt.\n"
            "(3) Escrow & Refund: Held in escrow pursuant to Section 551 BGB. Refunded within 30 days after checkout protocol less verified claims. Setting off deposit against ongoing rent is prohibited."
        ),
        (
            "§ 4 Zahlungsbedingungen, Fälligkeit & Verzug\n",
            f"(1) Fälligkeit: Miete ist monatlich im Voraus bis spätestens zum 3. Werktag kostenfrei zu überweisen.\n"
            f"(2) Bankverbindung: Inhaber: {data['RENT_ACCOUNT_HOLDER']} | IBAN: {data['RENT_IBAN']} | BIC: {data['RENT_BIC']}.\n"
            "(3) Verzug: Mahngebühr von 5,00 € je Mahnung. Rückstand von zwei Monatsmieten berechtigt zur fristlosen Kündigung gem. § 543 Abs. 2 BGB.",
            "§ 4 Payment Terms, Due Dates & Late Fees\n",
            "(1) Due Date: Rent is due in advance on or before the 3rd business day of each month via wire transfer.\n"
            f"(2) Account: Holder: {data['RENT_ACCOUNT_HOLDER']} | IBAN: {data['RENT_IBAN']} | BIC: {data['RENT_BIC']}.\n"
            "(3) Default: Overdue reminder fee of €5.00 per notice. Rent arrears of two months permit immediate cancellation under § 543 (2) BGB."
        ),
        (
            "§ 5 Zusammenleben in der WG, Lärm & Ruhezeiten\n",
            "(1) Gegenseitige Rücksichtnahme: Die WG regelt das Zusammenleben partnerschaftlich und eigenverantwortlich.\n"
            "(2) Ruhezeiten: Gesetzliche Ruhezeiten sind strikt einzuhalten: 13:00–15:00 Uhr sowie 22:00–08:00 Uhr. An Sonn- und Feiertagen absolute Ruhe. Audiogeräte nur auf Zimmerlautstärke.\n"
            "(3) Partys: Störende Partys sind verboten. Geselliges Beisammensein ist ab 22:00 Uhr auf Zimmerlautstärke zu senken.",
            "§ 5 Flatshare Cohabitation, Quiet Hours & Noise Control\n",
            "(1) Mutual Courtesy: Flatmates manage everyday co-living respectfully and independently.\n"
            "(2) Quiet Hours: Mandatory quiet hours: 1:00pm–3:00pm and 10:00pm–8:00am. Full silence on Sundays and public holidays. Media devices at room volume only.\n"
            "(3) Parties: Disruptive parties are prohibited. Gatherings must be quiet by 10:00pm."
        ),
        (
            "§ 6 Wohnungsnutzung, Lüften & Heizen\n",
            "(1) Wohnzweck: Ausschließlich zu privaten Wohnzwecken gestattet. Keine gewerbliche Untervermietung.\n"
            "(2) Lüften & Schimmelschutz: Zur Vermeidung von Schimmel ist die Wohnung ganzjährig mind. 2x täglich für 5–10 Min. per Stoßlüften zu lüften. Dauerkippstellung im Winter ist unzulässig.\n"
            "(3) Heizen: Thermostate beim Lüften herunterdrehen; in der Heizperiode ausreichende Mindesttemperatur halten.",
            "§ 6 Apartment Care, Ventilation & Heating\n",
            "(1) Residential Use: Exclusively for living purposes. Commercial activity or short-term subletting prohibited.\n"
            "(2) Ventilation: To prevent mold, cross-ventilate windows fully 2x daily for 5–10 minutes. Leaving windows tilted continuously in winter is forbidden.\n"
            "(3) Heating: Turn down thermostats during airing; keep adequate baseline warmth in cold months."
        ),
        (
            "§ 7 Rauchverbot, Drogenverbot & Brandschutz\n",
            "(1) Rauchverbot: Absolutes Rauchverbot in allen Innenräumen, Fluren und auf den Balkonen (inkl. E-Zigaretten).\n"
            "(2) Drogenverbot: Besitz, Konsum oder Handel illegaler Betäubungsmittel führt zur sofortigen Anzeige und fristlosen Kündigung.\n"
            "(3) Brandschutz: Offenes Feuer, unbeaufsichtigte Kerzen und Grillen auf dem Balkon sind untersagt.",
            "§ 7 Strict Non-Smoking, Drug Ban & Fire Safety\n",
            "(1) Non-Smoking: Strictly forbidden throughout all interior rooms, corridors, and balconies (including vapes).\n"
            "(2) Narcotics Ban: Illegal substances result in immediate police report and termination without notice.\n"
            "(3) Fire Safety: Open flames, unattended candles, and balcony barbecues are forbidden."
        ),
        (
            "§ 8 Haustiere & Beherbergung von Gästen\n",
            "(1) Haustiere: Die Haltung von Haustieren jeglicher Art ist nicht gestattet.\n"
            "(2) Gäste: Übernachtungsbesuch ist nur bei Anwesenheit des Mieters und maximal 5 Tage pro Kalendermonat zulässig. Dauerhafte Aufnahme dritter Personen gilt als unzulässige Gebrauchsüberlassung.",
            "§ 8 Pets Prohibition & Overnight Guest Policy\n",
            "(1) Pets: Keeping pets of any kind is prohibited.\n"
            "(2) Guests: Overnight visitors allowed only while tenant is present, capped at maximum 5 days per calendar month. Permanent third-party accommodation is unauthorized subletting."
        ),
        (
            "§ 9 Internetnutzung, Haftungsfreistellung & Urheberrecht\n",
            "(1) Bereitstellung: Highspeed-WLAN wird bereitgestellt. Mieter trägt alleinige zivil- und strafrechtliche Verantwortung für jegliche Online-Aktivitäten.\n"
            "(2) Urheberrecht: Peer-to-Peer, BitTorrent und illegale Downloads sind strengstens verboten.\n"
            "(3) Freistellung: Mieter stellt Vermieter von Abmahn- und Anwaltskosten vollumfänglich frei. Bei Missbrauch erfolgt sofortige Internetsperre und Weitergabe der Mieterdaten an Behörden.",
            "§ 9 Internet Policy, Torrent Ban & Indemnification\n",
            "(1) Access: Highspeed Wi-Fi provided. Tenant assumes sole civil and criminal liability for all online traffic.\n"
            "(2) Copyright Ban: BitTorrent, P2P file-sharing, and illegal downloads are strictly prohibited.\n"
            "(3) Indemnification: Tenant fully indemnifies landlord against legal claims. Access is revoked upon violation and logs forwarded to police."
        ),
        (
            "§ 10 Obhutspflichten, Wände & Mülltrennung\n",
            "(1) Wände: Wandbohrungen, Dübel, Nägel und klebende Folien, die den Putz beschädigen, sind untersagt.\n"
            "(2) Abflüsse: Keine Speisereste, Fette, Öle oder Hygieneartikel in Spülbecken, Dusche oder WC entsorgen.\n"
            "(3) Mülltrennung: Konsequente Mülltrennung gem. Berliner Stadtreinigung (BSR) (Gelbe Tonne, Papier, Bio, Restmüll).",
            "§ 10 Duty of Care, Wall Treatment & Waste Sorting\n",
            "(1) Walls: Drilling holes, inserting wall plugs/nails, or adhesive tapes damaging plaster is prohibited.\n"
            "(2) Plumbing: No food scraps, cooking grease, oils, or sanitary items down kitchen or bathroom drains.\n"
            "(3) Waste Sorting: Systematic recycling per Berlin BSR standards (packaging, paper, bio, residual)."
        ),
        (
            "§ 11 Übergabe und Rückgabe der Mietsache\n",
            "(1) Einzug: Übergabe erfolgt nach Geldeingang von Kaution und 1. Miete anhand eines detaillierten Übergabeprotokolls.\n"
            "(2) Auszug: Rückgabe am letzten Vertragstag bis spätestens 11:00 Uhr im geräumten, besenreinen Zustand samt aller Schlüssel.\n"
            "(3) Hinterlassene Sachen: Zurückgelassene Gegenstände werden auf Mieterkosten entsorgt oder eingelagert.",
            "§ 11 Handover Protocol & Return of Premises\n",
            "(1) Move-In: Keys transferred after full receipt of deposit and 1st rent via signed handover protocol.\n"
            "(2) Move-Out: Room returned swept clean, fully cleared, and with all keys by 11:00am on final lease day.\n"
            "(3) Left Belongings: Disposed of or stored at tenant's expense."
        ),
        (
            "§ 12 Schäden, Kleinreparaturen & Schlüsselersatz\n",
            "(1) Mängelmeldung: Schäden sind dem Vermieter unverzüglich anzuzeigen.\n"
            "(2) Kleinreparaturen: Kleinere Instandhaltungen bis 100,00 € im Einzelfall (max. 8% der Jahresmiete bzw. 350,00 € jährlich) trägt der Mieter.\n"
            "(3) Schlüsselverlust-Pauschalen: Haustür: 680,00 € | Wohnung: 120,00 € | Zimmer: 90,00 € | Briefkasten: 70,00 € zzgl. ggf. Austausch der Zentralschließanlage.",
            "§ 12 Damages, Minor Repairs & Lost Keys Schedule\n",
            "(1) Defect Notice: Damages must be reported immediately.\n"
            "(2) Minor Repairs: Repairs on fixtures up to €100.00 per instance (annual cap: 8% of rent or €350.00) borne by tenant.\n"
            "(3) Lost Keys Schedule: Building door: €680.00 | Apartment: €120.00 | Room: €90.00 | Mailbox: €70.00 plus master lock system replacement if required."
        ),
        (
            "§ 13 Betretungsrecht des Vermieters\n",
            "(1) Vorankündigung: Zutritt werktags (09:00–12:00 und 15:00–20:00 Uhr) nach angemessener Ankündigung (24–48 Std.) zur Prüfung des Zustands oder für Handwerker.\n"
            "(2) Nachmieterbesichtigungen: In den letzten 4 Wochen der Mietzeit hat der Mieter bis zu 5 Besichtigungstermine nach Vorankündigung zu dulden.\n"
            "(3) Notfälle: Bei Gefahr im Verzug ist der Zutritt jederzeit ohne Vorankündigung gestattet.",
            "§ 13 Landlord Right of Inspection & Access\n",
            "(1) Advance Notice: Access on workdays (9:00am–12:00pm & 3:00pm–8:00pm) upon 24–48h notice for inspections or maintenance.\n"
            "(2) Viewings: Tenant permits up to 5 prospective tenant viewings during the final 4 weeks of stay.\n"
            "(3) Emergencies: Immediate access permitted anytime in case of imminent danger."
        ),
        (
            "§ 14 Kündigung, Vorzeitige Entlassung, Schriftform & Salvatorische Klausel\n",
            "(1) Ausschluss ordentlicher Kündigung: Während der festen Mietdauer ist die ordentliche Kündigung ausgeschlossen.\n"
            "(2) Nachmieterstellung: Vorzeitige Entlassung mit 3 Monaten Frist zum Monatsende bei Gestellung eines akzeptierten, solventen Nachmieters.\n"
            "(3) Fristlose Kündigung (§ 543 BGB): Zulässig bei Verzug von 2 Monatsmieten, schwerer Pflichtverletzung, Drogen, Reinigungssäumung oder Erhalt von 2 Abmahnungen.\n"
            "(4) Schriftform & Salvatorische Klausel: Mündliche Nebenabreden ungültig. Änderungen nur schriftlich. Sollten Klauseln unwirksam sein, bleibt Restvertrag gültig. Erfüllungsort Berlin.",
            "§ 14 Termination, Replacement Tenant, Written Form & Severability\n",
            "(1) No Ordinary Notice: Ordinary early termination excluded during the fixed lease.\n"
            "(2) Replacement Tenant: Early release with 3 months' notice to month-end upon providing an accepted solvent successor tenant.\n"
            "(3) Extraordinary Cancellation (§ 543 BGB): Applicable for 2 months rent arrears, severe contract breach, drugs, cleaning neglect, or 2 written warnings.\n"
            "(4) Written Form & Severability: No verbal side agreements. Invalid provisions do not invalidate agreement. Jurisdiction Berlin."
        )
    ]

    t_agb = doc.add_table(rows=len(clauses_data), cols=2)
    t_agb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (title_de, text_de, title_en, text_en) in enumerate(clauses_data):
        row = t_agb.rows[i]
        row.cells[0].width = Inches(3.6)
        row.cells[1].width = Inches(3.6)

        p_de = row.cells[0].paragraphs[0]
        r_de = p_de.add_run(title_de)
        r_de.font.bold = True
        r_de.font.size = Pt(9.5)
        r_de_txt = p_de.add_run(text_de)
        r_de_txt.font.size = Pt(8.5)

        p_en = row.cells[1].paragraphs[0]
        r_en = p_en.add_run(title_en)
        r_en.font.bold = True
        r_en.font.size = Pt(9.5)
        r_en.font.color.rgb = RGBColor(71, 85, 105)
        r_en_txt = p_en.add_run(text_en)
        r_en_txt.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(25)

    # Signatures Table (NO timestamps)
    t_sig = doc.add_table(rows=1, cols=2)
    t_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r in t_sig.rows:
        r.cells[0].width = Inches(3.6)
        r.cells[1].width = Inches(3.6)

    t_sig.rows[0].cells[0].paragraphs[0].add_run(f"_____________________________________\n{data['LANDLORD_SIGN_NAME']}\nVermieter / Landlord\nDatum / Date: ____________________")
    t_sig.rows[0].cells[1].paragraphs[0].add_run(f"_____________________________________\n{data['TENANT_SIGN_NAME']}\nMieter / Tenant\nDatum / Date: ____________________")

    doc.save(output_path)
    return output_path

def convert_html_to_pdf(html_path, pdf_path):
    chrome_bin = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    if not os.path.exists(chrome_bin):
        raise FileNotFoundError(f"Chrome not found at {chrome_bin}")

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

def main():
    print("Building official Homelyo Rental Contract Templates with comprehensive AGB...")
    os.makedirs(os.path.dirname(HTML_TEMPLATE_PATH), exist_ok=True)

    html = generate_html_template()
    with open(HTML_TEMPLATE_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✓ Master HTML template saved: {HTML_TEMPLATE_PATH}")

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

if __name__ == "__main__":
    main()
