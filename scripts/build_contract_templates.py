#!/usr/bin/env python3
"""
Homelyo Master Bilingual Rental Contract Builder (7-Page Format)
---------------------------------------------------------------
Replicates the exact format, typography, clause layout, text placement,
and pagination of the official signed agreement:
'/Users/jayashreesundarasamy/Downloads/SIGNED_CONTRACT_67da88a52ff5e11a1f9d8ffa_20250325. (1).pdf'

Total Pages: Exactly 7 Pages
- Page 1: Title ("Mietvertrag" / "Rental Agreement"), Parties (Landlord & Tenant in dual columns), Footer "1/7"
- Page 2: Header "2/7", Präambel / Preamble, 1. Mietsache / Rental property, 2. Mietdauer / Rental duration
- Page 3: Header "3/7", Continuation § 545 BGB, 3. Miete / Rent, 4. Mietzahlungen / Payment of rent
- Page 4: Header "4/7", Bank Details (Rent & Deposit), 5. Mietsicherheit/ Kaution / Security deposit
- Page 5: Header "5/7", 6. Kündigung / Termination, 7. Nutzung der Mietsache / Use of property, 8. Instandhaltungspflicht
- Page 6: Header "6/7", 9. Betreten der Mietsache, 10. Rückgabe, 11. Gerichtsstand, 12. Datenschutz, 13. Individuelle Vereinbarungen
- Page 7: Header "7/7", 14. Salvatorische Klausel / Severability clause, Signatures Block (Landlord & Tenant)

Google Drive Target Folders:
- Templates Folder ID: '16gimsgZ5BaVDk766sMDAuZWJcKJtKdqe'
- Generated Contracts Folder ID: '1CZcV4XWuY4pSpfNjkubdHsxprKg1f2xB'
"""

import os
import sys
import datetime
import subprocess
import urllib.request
import json
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

TEMPLATES_FOLDER_ID = "16gimsgZ5BaVDk766sMDAuZWJcKJtKdqe"
GENERATED_CONTRACTS_FOLDER_ID = "1CZcV4XWuY4pSpfNjkubdHsxprKg1f2xB"
GDRIVE_TEMPLATE_FILE_ID = "1Cj9E9nbcCbgV4HuT2ZLxdfNwiDLYVT5t"

HTML_TEMPLATE_PATH = "/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/templates/contract_template.html"
DOCX_TEMPLATE_PATH = "/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/template_contract.docx"
FINAL_PDF_PATH = "/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/contracts_generated/Rental_Contract_Master_Final.pdf"

ROOM_SPECS = {
    "KE1": {
        "title_de": "Master-Schlafzimmer",
        "title_en": "Master Bedroom",
        "size": "21.50",
        "floor": "2",
        "room_num": "1",
        "rent": 640.00,
        "deposit": 1000.00,
    },
    "KE2": {
        "title_de": "Sonnige Balkon-Suite",
        "title_en": "Private Balcony Suite",
        "size": "23.92",
        "floor": "2",
        "room_num": "2",
        "rent": 649.00,
        "deposit": 1000.00,
    },
    "KE3": {
        "title_de": "Ruhiges Gartenzimmer",
        "title_en": "Quiet Garden View",
        "size": "12.80",
        "floor": "2",
        "room_num": "3",
        "rent": 615.00,
        "deposit": 1000.00,
    },
    "KE4": {
        "title_de": "Studio-Zimmer",
        "title_en": "Cozy Studio Room",
        "size": "11.20",
        "floor": "2",
        "room_num": "4",
        "rent": 569.00,
        "deposit": 1000.00,
    }
}

def num_to_words_de(n):
    n = int(round(n))
    if n == 0:
        return "null"
    units = ["", "ein", "zwei", "drei", "vier", "fünf", "sechs", "sieben", "acht", "neun"]
    units_exact = ["", "eins", "zwei", "drei", "vier", "fünf", "sechs", "sieben", "acht", "neun"]
    teens = ["zehn", "elf", "zwölf", "dreizehn", "vierzehn", "fünfzehn", "sechzehn", "siebzehn", "achtzehn", "neunzehn"]
    tens = ["", "", "zwanzig", "dreißig", "vierzig", "fünfzig", "sechzig", "siebzig", "achtzig", "neunzig"]

    def _under_thousand(val):
        res = ""
        h = val // 100
        rem = val % 100
        if h == 1:
            res += "einhundert"
        elif h > 1:
            res += units[h] + "hundert"
        if rem == 0:
            return res
        if rem < 10:
            res += units_exact[rem] if not res else units[rem]
        elif rem < 20:
            res += teens[rem - 10]
        else:
            u = rem % 10
            t = rem // 10
            if u == 0:
                res += tens[t]
            else:
                res += units[u] + "und" + tens[t]
        return res

    words = ""
    thousands = n // 1000
    rem = n % 1000
    if thousands == 1:
        words += "eintausend"
    elif thousands > 1:
        words += _under_thousand(thousands) + "tausend"
    if rem > 0:
        words += _under_thousand(rem)
    return words

def num_to_words_en(n):
    n = int(round(n))
    if n == 0:
        return "zero"
    units = ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine"]
    teens = ["ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"]
    tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]

    def _under_thousand(val):
        parts = []
        h = val // 100
        rem = val % 100
        if h > 0:
            parts.append(units[h] + " hundred")
        if rem > 0:
            if parts:
                parts.append("and")
            if rem < 10:
                parts.append(units[rem])
            elif rem < 20:
                parts.append(teens[rem - 10])
            else:
                u = rem % 10
                t = rem // 10
                if u == 0:
                    parts.append(tens[t])
                else:
                    parts.append(tens[t] + "-" + units[u])
        return " ".join(parts)

    parts = []
    thousands = n // 1000
    rem = n % 1000
    if thousands > 0:
        parts.append(_under_thousand(thousands) + " thousand")
    if rem > 0:
        parts.append(_under_thousand(rem))
    return " ".join(parts)

def format_date_bilingual(date_str):
    if not date_str:
        return "", ""
    dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    de_months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
    en_months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    day = dt.day
    year = dt.year
    m_idx = dt.month - 1
    return f"{day} {de_months[m_idx]}, {year}", f"{day} {en_months[m_idx]}, {year}"

def render_contract_data(
    tenant_name,
    tenant_email="",
    tenant_phone="",
    room_code="KE2",
    start_date="2025-03-25",
    end_date="2025-09-30",
    rent_amount=None,
    deposit_amount=None,
    tenant_address_de=None,
    tenant_address_en=None,
    reason_de=None,
    reason_en=None,
    landlord_sign_date="20 Mar 2025, 10:39 GMT",
    tenant_sign_date="21 Mar 2025, 12:36 GMT",
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

    if not tenant_address_de:
        tenant_address_de = "Type 4 building no 3 flat no 4 esis hospital quarters , Vashi sector-5 sagar vihar road, 400703 Navi Mumbai, Indien"
    if not tenant_address_en:
        tenant_address_en = "Type 4 building no 3 flat no 4 esis hospital quarters , Vashi sector-5 sagar vihar road, 400703 Navi Mumbai, India"

    if not reason_de:
        reason_de = "Aufgrund eines Studiensemesters außerhalb seines eigentlichen Studienortes."
    if not reason_en:
        reason_en = "Due to a semester of study outside of his/her normal place of study."

    default_indiv_de = "5 % Erhöhung der Miete pro Jahr<br>Wenn Sie es mit anderen Personen teilen, müssen Sie den Kauf gemeinsamer Gegenstände wie WC-Reinigungsflüssigkeit unter den Mietern einplanen."
    default_indiv_en = "5% increase on rent/year<br>if shared with other people, need to plan buying common items like wc cleaning liquid among tenants."

    data = {
        "TENANT_NAME": tenant_name,
        "TENANT_ADDRESS_DE": tenant_address_de.replace("\n", "<br>"),
        "TENANT_ADDRESS_EN": tenant_address_en.replace("\n", "<br>"),
        "ROOM_CODE": room_code,
        "ROOM_SIZE": spec.get("size", "23.92"),
        "ROOM_FLOOR": spec.get("floor", "2"),
        "ROOM_TITLE_DE": spec.get("title_de", "Sonnige Balkon-Suite"),
        "ROOM_TITLE_EN": spec.get("title_en", "Private Balcony Suite"),
        "START_DATE_DE": start_de,
        "START_DATE_EN": start_en,
        "END_DATE_DE": end_de,
        "END_DATE_EN": end_en,
        "REASON_DE": reason_de,
        "REASON_EN": reason_en,
        "RENT_AMOUNT": rent_str,
        "RENT_WORDS_DE": rent_words_de,
        "RENT_WORDS_EN": rent_words_en,
        "FIRST_RENT_AMOUNT": rent_str,
        "CLEANING_FEE": "50",
        "LATE_FEE": "2,50",
        "DEPOSIT_AMOUNT": deposit_str,
        "DEPOSIT_WORDS_DE": deposit_words_de,
        "DEPOSIT_WORDS_EN": deposit_words_en,
        "INDIVIDUAL_AGREEMENTS_DE": (individual_de or default_indiv_de).replace("\n", "<br>"),
        "INDIVIDUAL_AGREEMENTS_EN": (individual_en or default_indiv_en).replace("\n", "<br>"),
        "LANDLORD_SIGN_DATE": landlord_sign_date,
        "TENANT_SIGN_DATE": tenant_sign_date
    }
    return data

def generate_html_template():
    return """<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<title>Mietvertrag / Rental Agreement - Homelyo</title>
<style>
@page {
  size: A4 portrait;
  margin: 14mm 15mm 14mm 15mm;
}
* {
  box-sizing: border-box;
}
body {
  font-family: Arial, Helvetica, "Liberation Sans", sans-serif;
  font-size: 8.7pt;
  line-height: 1.34;
  color: #111;
  margin: 0;
  padding: 0;
  -webkit-font-smoothing: antialiased;
}
.page {
  height: 269mm;
  max-height: 269mm;
  position: relative;
  page-break-after: always;
  overflow: hidden;
}
.page:last-child {
  page-break-after: avoid;
}
.page-footer-num {
  position: absolute;
  bottom: 0;
  right: 0;
  font-size: 8.5pt;
  color: #444;
}
.cols {
  display: flex;
  justify-content: space-between;
  width: 100%;
}
.col-de, .col-en {
  width: 48.2%;
}
.clause-h {
  font-size: 9.8pt;
  font-weight: bold;
  margin: 8px 0 3px 0;
  color: #000;
}
.clause-p {
  margin: 0 0 5px 0;
  text-align: justify;
}
.bank-grid {
  width: 100%;
  margin: 4px 0 6px 0;
  font-size: 8.4pt;
}
.bank-grid td {
  padding: 1px 0;
  vertical-align: top;
}
.bank-lbl {
  color: #444;
  font-size: 8pt;
}
.bank-val {
  font-weight: bold;
}
.sub-list {
  padding-left: 14px;
  margin: 2px 0 4px 0;
}
.sub-list li {
  margin-bottom: 2px;
}
.sign-table {
  width: 100%;
  margin-top: 40px;
}
.sign-table td {
  width: 50%;
  vertical-align: top;
  font-size: 9.2pt;
  line-height: 1.45;
}
</style>
</head>
<body>

<!-- ================================= PAGE 1 ================================= -->
<div class="page" id="page-1">
  <div style="text-align: center; margin-top: 60px; margin-bottom: 80px;">
    <div style="font-size: 27pt; font-weight: bold; color: #000; letter-spacing: -0.5px;">Mietvertrag</div>
    <div style="font-size: 17pt; color: #222; margin-top: 8px;">Rental Agreement</div>
  </div>

  <div class="cols" style="margin-bottom: 70px;">
    <div class="col-de">
      <div style="font-weight: bold; font-size: 11pt; margin-bottom: 26px;">Zwischen Vermieter:</div>
      <div style="font-size: 8.5pt; color: #555;">Vorname, Nachname</div>
      <div style="font-size: 10.5pt; font-weight: bold; margin-bottom: 22px;">Kylasam NA</div>
      <div style="font-size: 8.5pt; color: #555;">Anschrift</div>
      <div style="font-size: 10pt; line-height: 1.4;">Forckenbeckstraße 64,<br>14199 Berlin, Deutschland</div>
    </div>
    <div class="col-en">
      <div style="font-weight: bold; font-size: 11pt; margin-bottom: 26px;">The landlord:</div>
      <div style="font-size: 8.5pt; color: #555;">Name, Surname</div>
      <div style="font-size: 10.5pt; font-weight: bold; margin-bottom: 22px;">Kylasam NA</div>
      <div style="font-size: 8.5pt; color: #555;">Address</div>
      <div style="font-size: 10pt; line-height: 1.4;">Forckenbeckstraße 64,<br>14199 Berlin, Germany</div>
    </div>
  </div>

  <div class="cols">
    <div class="col-de">
      <div style="font-weight: bold; font-size: 11pt; margin-bottom: 26px;">Zwischen Mieter:</div>
      <div style="font-size: 8.5pt; color: #555;">Vorname, Nachname</div>
      <div style="font-size: 10.5pt; font-weight: bold; margin-bottom: 22px;">{{TENANT_NAME}}</div>
      <div style="font-size: 8.5pt; color: #555;">Anschrift</div>
      <div style="font-size: 10pt; line-height: 1.4;">{{TENANT_ADDRESS_DE}}</div>
    </div>
    <div class="col-en">
      <div style="font-weight: bold; font-size: 11pt; margin-bottom: 26px;">The tenant:</div>
      <div style="font-size: 8.5pt; color: #555;">Name, Surname</div>
      <div style="font-size: 10.5pt; font-weight: bold; margin-bottom: 22px;">{{TENANT_NAME}}</div>
      <div style="font-size: 8.5pt; color: #555;">Address</div>
      <div style="font-size: 10pt; line-height: 1.4;">{{TENANT_ADDRESS_EN}}</div>
    </div>
  </div>

  <div class="page-footer-num">1/7</div>
</div>

<!-- ================================= PAGE 2 ================================= -->
<div class="page" id="page-2">
  <div class="cols">
    <div class="col-de">
      <div class="clause-h" style="margin-top:0;">Präambel</div>
      <p class="clause-p">Der Vermieter ist nicht zwingend Eigentümer des Mietobjektes, sondern hat den Mietgegenstand gegebenenfalls vom Eigentümer oder sonstigen Berechtigten zum Zwecke der Weitervermietung angemietet. Der Vermieter vermietet möblierten Wohnraum ausschließlich für begrenzte Zeit zum vorübergehenden Gebrauch.</p>

      <div class="clause-h">1. Mietsache</div>
      <p class="clause-p">(1) Der Vermieter vermietet dem Mieter zur Nutzung durch folgende Person(en)<br>
      <strong>1. {{TENANT_NAME}}</strong><br>
      zum vorübergehenden Gebrauch folgende Wohnung im Hause <strong>Lichtenrader Damm 80, 12305 Berlin, Deutschland</strong> , im {{ROOM_FLOOR}} Geschoss bestehend aus 1 Zimmer(n).<br>
      Die Wohnfläche beträgt ca. <strong>{{ROOM_SIZE}} m²</strong>. Diese Angabe dient wegen möglicher Messfehler nicht zur Festlegung des Mietgegenstandes. Der räumliche Umfang der gemieteten Sache ergibt sich vielmehr aus der Anzahl der vermieteten Räume.</p>
      <p class="clause-p">(2) Das Mietobjekt ist voll möbliert und komplett eingerichtet.</p>
      <p class="clause-p">(3) Bei Wohnungsübergabe werden dem Mieter folgende Schlüssel ausgehändigt:<br>
      <em>Wie bei Übergabe vereinbart</em></p>

      <div class="clause-h">2. Mietdauer</div>
      <p class="clause-p">(1) Der Mietvertrag beginnt am <strong>{{START_DATE_DE}}</strong> und endet am <strong>{{END_DATE_DE}}</strong>.</p>
      <p class="clause-p">(2) Die Wohnung ist gemäß § 549 Abs. 2 Nr. 1 BGB lediglich zum vorübergehenden Gebrauch vermietet. Der Vermieter vermietet das Mietobjekt ausschließlich im Rahmen des vorübergehenden Gebrauchs. Auf Seiten des Mieters besteht ein vorübergehender Sonderbedarf, welcher durch die Anmietung der voll möblierten Mietsache im Rahmen des § 549 Abs. 2 Nr. 1 BGB gedeckt werden soll. Der Mieter bestätigt einen vorübergehenden Sonderbedarf <strong>{{REASON_DE}}</strong></p>
      <p class="clause-p">Sofern der vorübergehende Wohnbedarf aufgrund geänderter Umstände über den vereinbarten Zeitraum hinaus fortbesteht, kommt eine Verlängerung des Mietverhältnisses in Betracht. Hierfür hat der Mieter dem Vermieter den Grund zum vorübergehenden Gebrauch spätestens einen vollen Monat vor Mietende mitzuteilen. Der Vermieter ist nicht verpflichtet, den neuen oder aktualisierten Grund für den Aufenthalt zu akzeptieren und das Mietverhältnis fortzusetzen.</p>
      <p class="clause-p">(3) Setzt der Mieter nach Ablauf der Mietzeit den Gebrauch der Mietsache fort, so verlängert sich das</p>
    </div>

    <div class="col-en">
      <div class="clause-h" style="margin-top:0;">Preamble</div>
      <p class="clause-p">The landlord is not necessarily the owner of the rental property but has rented the property from the owner or another authorised party for the purpose of subletting it. The landlord rents a furnished living space exclusively for a limited period of time for temporary use.</p>

      <div class="clause-h">1. Rental property</div>
      <p class="clause-p">(1) The landlord rents to the tenant for temporary use on behalf of the following person(s)<br>
      <strong>1. {{TENANT_NAME}}</strong><br>
      the following apartment is located in <strong>Lichtenrader Damm 80, 12305 Berlin, Germany</strong> , on {{ROOM_FLOOR}} floor consisting of 1 rooms.<br>
      The residential space consists of approximately <strong>{{ROOM_SIZE}} m²</strong>. Due to possible measurement errors, this information is not conducive to define the rental property. Instead, the size of the rental property is determined by the number of rented rooms.</p>
      <p class="clause-p">(2) The rental property is entirely furnished and completely equipped.</p>
      <p class="clause-p">(3) When the apartment is handed over, the following keys are provided to the tenant:<br>
      <em>As agreed during key handover</em></p>

      <div class="clause-h">2. Rental duration</div>
      <p class="clause-p">(1) The rental contract starts on <strong>{{START_DATE_EN}}</strong> and ends on <strong>{{END_DATE_EN}}</strong>.</p>
      <p class="clause-p">(2) The apartment is rented for temporary use only, in accordance with section 549, paragraph 2(1) of the German Civil Code (Bürgerliches Gesetzbuch, BGB), the landlord rents the rental property exclusively for temporary use. The tenant has a special need for temporary accommodation, to be covered by renting the fully furnished rental property pursuant to section 549, paragraph 2(1) of the BGB. The tenant confirms a temporary special need <strong>{{REASON_EN}}</strong></p>
      <p class="clause-p">If the temporary housing need continues beyond the agreed period due to changed circumstances, an extension of the rental period is possible. For this purpose, the tenant must inform the landlord of the reason for temporary use no later than one full month before the end of the rental period. The landlord is not obliged to accept the new or updated reason for residency and to resume the rental agreement.</p>
      <p class="clause-p">(3) If the tenant continues to use the rental property after the termination of the rental period, the rental relationship shall not be extended indefinitely. Section 545 of the BGB is excluded.</p>
    </div>
  </div>

  <div class="page-footer-num">2/7</div>
</div>

<!-- ================================= PAGE 3 ================================= -->
<div class="page" id="page-3">
  <div class="cols">
    <div class="col-de">
      <p class="clause-p">Mietverhältnis nicht auf unbestimmte Zeit. § 545 BGB wird ausgeschlossen.</p>

      <div class="clause-h">3. Miete</div>
      <p class="clause-p">(1) Die Miete beträgt monatlich <strong>{{RENT_AMOUNT}}€</strong>, in Worten: <strong>{{RENT_WORDS_DE}} EURO</strong>.<br>
      Hierin enthalten ist<br>
      <span style="display:inline-block; margin-left:10px;">• eine Nettokaltmiete inklusive Möblierungszuschlag</span><br>
      <span style="display:inline-block; margin-left:10px;">• Nebenkosten inkl. Betriebs-, Heiz-, Warmwasser-, Strom- und Internetkosten</span></p>

      <p class="clause-p">(2) Der Mieter hat für eine sachgemäße, nicht übertriebene Nutzung (Fair Usage) der Nebenkosten zu sorgen. Hier ist besonders auf den sachgemäßen Verbrauch von Strom, Wasser und Heizung zu achten.</p>
      <p class="clause-p">Die Gesamtmiete beinhaltet eine Internet-Flatrate. Weitere Kosten, die durch die Nutzung der Internetverbindung entstehen (z.B. kostenpflichtige oder illegale Downloads etc.), sind vom Mieter der Wohnung selbst zu tragen und werden mit dem Vermieter abgerechnet. Beim Einzug stellt der Vermieter dem Mieter den Router und die Zugangsdaten für die WLAN-Verbindung zur Verfügung. Beide Parteien sind sich einig, dass der Mieter allein für die Nutzung der Internetverbindungen und die illegale Internetnutzung verantwortlich ist. Bei illegaler Nutzung kann der Vermieter die Internetverbindung ohne Vorankündigung sperren. In diesem Fall hat der Mieter keinen Anspruch auf Mietminderung.</p>
      <p class="clause-p">Weitere, unmittelbar mit der Wohnung zusammenhängende Verbrauchskosten, z. B. Telefonkosten, trägt der Mieter selbst.</p>
      <p class="clause-p">(3) Bei einem Mietverhältnis mit einer Dauer von unter 6 Monaten findet der aktuell gesetzlich gültige Umsatzsteuersatz Anwendung.</p>
      <p class="clause-p">(4) Der Mieter hat zusätzlich zur laufenden Miete eine Endreinigungspauschale in Höhe von <strong>{{CLEANING_FEE}}€</strong> zu leisten, welche bei Rückgabe der Wohnung bar zu übergeben oder zu überweisen ist oder nach Absprache mit dem Vermieter von der Kaution abgezogen wird.</p>

      <div class="clause-h">4. Mietzahlungen</div>
      <p class="clause-p">(1) Die erste Miete ist bei Vertragsunterzeichnung fällig und beläuft sich auf <strong>€{{FIRST_RENT_AMOUNT}}</strong>. Leistet der Mieter die vereinbarte Miete nicht, besteht kein Anspruch auf Übergabe der Mietsache.</p>
      <p class="clause-p">(2) Im Folgenden ist die Miete monatlich im Voraus, spätestens bis zum 1. Werktag des jeweiligen Monats, kostenfrei an den Vermieter auf folgendes Konto zu leisten:</p>
    </div>

    <div class="col-en">
      <div style="height: 18px;"></div>
      <div class="clause-h">3. Rent</div>
      <p class="clause-p">(1) The rent amounts to <strong>€{{RENT_AMOUNT}}</strong> per month, in words: <strong>{{RENT_WORDS_EN}} EURO</strong>.<br>
      This includes<br>
      <span style="display:inline-block; margin-left:10px;">• basic rent (ger:, Nettokaltmiete) which includes a furniture surcharge</span><br>
      <span style="display:inline-block; margin-left:10px;">• Additional expenses lump sum incl. operating, utility, heating, electricity and internet cost</span></p>

      <p class="clause-p">(2) The tenant must ensure proper, non-exaggerated use (fair usage) of the utility costs in particular a proper consumption of electricity, water and heating is important.</p>
      <p class="clause-p">The total rent includes an Internet flat rate. Further costs arising from the use of the Internet connection (e.g. fee-based or illegal downloads, etc.) are to be borne by the tenant of the apartment and will be settled with the landlord. When moving in, the landlord provides the tenant with the router and the login details to connect to the WLAN. Both parties agree that the tenant is solely responsible for the Internet connection and liable for the illegal use of the Internet. In the event of illegal use, the landlord can block the Internet connection without prior notice. In this case, the tenant cannot claim a reduction in rent.</p>
      <p class="clause-p">Additional consumption costs directly related to the apartment (e.g. telephone costs) shall be borne by the tenant.</p>
      <p class="clause-p">(3) In the case of a rental agreement with a duration of less than 6 months, the currently valid statutory VAT rate shall apply.</p>
      <p class="clause-p">(4) In addition to the regular rent, the tenant must pay a fee for the final clean-up of <strong>€{{CLEANING_FEE}}</strong>, which must be paid in cash or by bank transfer at the time of handover of the apartment, or otherwise deducted from the deposit upon consultation with the landlord.</p>

      <div class="clause-h">4. Payment of rent</div>
      <p class="clause-p">(1) The first rent is due after contract signing and amounts to <strong>{{FIRST_RENT_AMOUNT}}€</strong>. If the tenant does not pay the agreed rent, there is no entitlement to occupy the rental property.</p>
      <p class="clause-p">(2) In the subsequent period, the rent is to be paid on a monthly basis in advance, at the latest by the first working day of the corresponding month, free of any charges to the landlord, to the following account:</p>
    </div>
  </div>

  <div class="page-footer-num">3/7</div>
</div>

<!-- ================================= PAGE 4 ================================= -->
<div class="page" id="page-4">
  <div class="cols">
    <div class="col-de">
      <table class="bank-grid">
        <tr><td class="bank-lbl">Kontoinhaber:</td></tr>
        <tr><td class="bank-val">Kylasam Nallagoundanpalayam Arruchamy</td></tr>
        <tr><td class="bank-lbl">IBAN:</td></tr>
        <tr><td class="bank-val">DE87100110012624542768</td></tr>
        <tr><td class="bank-lbl">BIC:</td></tr>
        <tr><td class="bank-val">NTSBDEB1XXX</td></tr>
        <tr><td class="bank-lbl">Verwendungszweck:</td></tr>
        <tr><td class="bank-val">{{TENANT_NAME}} "Monat / Jahr"</td></tr>
      </table>

      <p class="clause-p">(3) Für die Rechtzeitigkeit der Zahlung kommt es nicht auf die Absendung, sondern auf den Eingang des Geldes an.</p>
      <p class="clause-p">(4) Der Mieter hat für jede Mahnung/ Zahlungserinnerung wegen verspäteter Zahlung der Miete {{LATE_FEE}} Euro für pauschalierte Mahnkosten zu zahlen, sofern der Mieter nicht nachweist, dass dem Vermieter geringere bzw. keine Kosten entstanden sind.</p>

      <div class="clause-h">5. Mietsicherheit/ Kaution</div>
      <p class="clause-p">(1) Der Mieter leistet für die Erfüllung seiner Pflichten aus diesem Mietvertrag zu Beginn des Mietverhältnisses eine Sicherheit in Höhe von <strong>{{DEPOSIT_AMOUNT}}€</strong>, an den Vermieter auf folgendes Konto:</p>

      <table class="bank-grid">
        <tr><td class="bank-lbl">Kontoinhaber:</td></tr>
        <tr><td class="bank-val">Kylasam Nallagoundanpalayam Arruchamy</td></tr>
        <tr><td class="bank-lbl">IBAN:</td></tr>
        <tr><td class="bank-val">DE25120400000011466000</td></tr>
        <tr><td class="bank-lbl">BIC:</td></tr>
        <tr><td class="bank-val">COBADEFFXXX</td></tr>
        <tr><td class="bank-lbl">Verwendungszweck:</td></tr>
        <tr><td class="bank-val">{{TENANT_NAME}} "Kaution"</td></tr>
      </table>
      <p class="clause-p">oder in bar bei Mietvertragsschluss.</p>

      <p class="clause-p">(2) Leistet der Mieter die vereinbarte Mietsicherheit nicht, besteht kein Anspruch auf Übergabe der Mietsache.</p>
      <p class="clause-p">(3) Nach Vertragsende und Rückgabe der Mietsache rechnet der Vermieter die Kaution gegenüber dem Mieter ab und zahlt, wenn sich die Wohnung in dem Zustand befindet, in dem sie sich am Anfang befand (ausgenommen Verschleiß) den errechneten Kautionsbetrag aus. Ein Zurückbehaltungsrecht des Vermieters besteht nicht, solange keine begründeten Gegenansprüche dargelegt werden (z.B. Schäden an Möbeln, nicht bezahlte Miete).</p>
    </div>

    <div class="col-en">
      <table class="bank-grid">
        <tr><td class="bank-lbl">Account holder:</td></tr>
        <tr><td class="bank-val">Kylasam Nallagoundanpalayam Arruchamy</td></tr>
        <tr><td class="bank-lbl">IBAN:</td></tr>
        <tr><td class="bank-val">DE87100110012624542768</td></tr>
        <tr><td class="bank-lbl">BIC:</td></tr>
        <tr><td class="bank-val">NTSBDEB1XXX</td></tr>
        <tr><td class="bank-lbl">Payment reference:</td></tr>
        <tr><td class="bank-val">{{TENANT_NAME}} "Month / Year"</td></tr>
      </table>

      <p class="clause-p">(3) Punctual payment does not depend on the time of transfer, but on the receipt of the money.</p>
      <p class="clause-p">(4) For each reminder/overdue payment notice attributable to the late payment of the rent, the tenant must pay a fee of €{{LATE_FEE}} as a reminder cost, unless the tenant proves that the landlord has incurred lower or no expenses.</p>

      <div class="clause-h">5. Security deposit</div>
      <p class="clause-p">(1) For the fulfilment of his/ her obligations from this rental agreement, the tenant shall pay a security deposit of <strong>€{{DEPOSIT_AMOUNT}}</strong>, to the landlord due upon the beginning of rental period, to the following bank account:</p>

      <table class="bank-grid">
        <tr><td class="bank-lbl">Account holder:</td></tr>
        <tr><td class="bank-val">Kylasam Nallagoundanpalayam Arruchamy</td></tr>
        <tr><td class="bank-lbl">IBAN:</td></tr>
        <tr><td class="bank-val">DE25120400000011466000</td></tr>
        <tr><td class="bank-lbl">BIC:</td></tr>
        <tr><td class="bank-val">COBADEFFXXX</td></tr>
        <tr><td class="bank-lbl">Payment reference:</td></tr>
        <tr><td class="bank-val">{{TENANT_NAME}} "Deposit"</td></tr>
      </table>
      <p class="clause-p">or in cash upon conclusion of the rental contract.</p>

      <p class="clause-p">(2) If the tenant does not pay the agreed security deposit, there is no entitlement to occupy the rental property.</p>
      <p class="clause-p">(3) After termination of the contract and the return of the rental property, the landlord settles the security deposit with the tenant and pays back the deposit amount if the apartment is in the same condition as it was at the start of the contract (excluding wear and tear). The landlord cannot exercise a retention right as long as no justified counter-claims are presented (e.g. damage to furniture, unpaid rent).</p>
    </div>
  </div>

  <div class="page-footer-num">4/7</div>
</div>

<!-- ================================= PAGE 5 ================================= -->
<div class="page" id="page-5">
  <div class="cols">
    <div class="col-de">
      <div class="clause-h" style="margin-top:0;">6. Kündigung</div>
      <p class="clause-p">Der Mietvertrag wird für den in Ziffer 2 genannten Zeitraum fest abgeschlossen.<br>
      Das Recht der Parteien zur außerordentlichen Kündigung des Mietverhältnisses bleibt hiervon unberührt.</p>

      <div class="clause-h">7. Nutzung der Mietsache</div>
      <p class="clause-p">(1) Eine weitere Untervermietung, sonstige Gebrauchsüberlassung an Dritte oder gewerbliche Nutzung ist dem Mieter ausdrücklich nicht gestattet.<br>
      (2) Haustiere sind in der Wohnung nicht erlaubt<br>
      (3) Rauchen ist in der Wohnung nicht erlaubt</p>

      <div class="clause-h">8. Instandhaltungspflicht/ Bauliche Veränderungen</div>
      <p class="clause-p">(1) Der Mieter ist verpflichtet, die Mietsache sowie die überlassenen Möbel, Einrichtungen und Anlagen schonend und pfleglich zu behandeln. Er hat für eine ordnungsgemäße Reinigung der Mietsache und für eine ausreichende Lüftung und Beheizung des ihm überlassenen Hauses zu sorgen.</p>
      <p class="clause-p">(2) Zeigt sich ein wesentlicher Mangel der Mietsache oder wird eine Vorkehrung zum Schutz der Mietsache oder des Anwesens gegen eine nicht vorhergesehene Gefahr erforderlich, so hat der Mieter dies dem Vermieter unverzüglich mitzuteilen und ihm Gelegenheit zu geben, innerhalb angemessener Frist für Abhilfe zu sorgen. Der Mieter hat keinen Anspruch auf Ersatz seiner Aufwendungen, wenn er nicht dem Vermieter vorher diese Gelegenheit gegeben hat, es sei denn, es bestand Gefahr im Verzug.</p>
      <p class="clause-p">(3) Den Mieter trifft die Beweislast für den Umstand, dass er einen Schaden nicht zu vertreten hat, wenn die Schadensursache in einem Bereich der Mietsache liegt, der nur ihm als Mietpartei zugänglich ist.</p>
      <p class="clause-p">(4) Hat der Mieter einen Schaden zu vertreten, ist er verpflichtet, diesen sofort zu beseitigen. Das bedeutet eine angemessene Schadensersatzzahlung oder der Austausch des beschädigten Inventars.</p>
      <p class="clause-p">(5) Der Mieter ist nicht berechtigt, an den gemieteten Räumen bauliche Veränderungen vorzunehmen.</p>
    </div>

    <div class="col-en">
      <div class="clause-h" style="margin-top:0;">6. Termination</div>
      <p class="clause-p">The rental contract is concluded for the period specified in Section 2.<br>
      The right of the parties to an exceptional termination of the rent remains unaffected.</p>

      <div class="clause-h">7. Use of the rental property</div>
      <p class="clause-p">(1) A further subletting, other transfer of use to third parties or commercial use on behalf of the tenant is expressly prohibited.<br>
      (2) Pets are not allowed in the apartment<br>
      (3) Smoking in the apartment is not allowed</p>

      <div class="clause-h">8. Obligation to undertake maintenance/ structural modifications</div>
      <p class="clause-p">(1) The tenant is obliged to treat the rental property, as well as the furniture, furnishings and equipment with respect and care. The tenant shall ensure proper cleaning and sufficient ventilation and heating of the rental property.</p>
      <p class="clause-p">(2) If a substantial defect of the rental property becomes apparent or if a precautionary measure is necessary to protect the rental property or the premises against an unforeseen hazard, the tenant shall immediately inform the landlord and grant the opportunity to remedy the situation within a reasonable period of time. The tenant is not entitled to compensation for his/her expenses if no notice was given to the landlord concerning this opportunity, unless an imminent danger was present.</p>
      <p class="clause-p">(3) The burden of proof, stating that he/she is not responsible for damage, is borne by the tenant if the cause of the damage is located in an area of the rental property that is solely accessible to him/her as a rental party.</p>
      <p class="clause-p">(4) If the tenant is responsible for damage, he/she is obliged to remedy it immediately. This means an appropriate compensation or the replacement of the damaged inventory.</p>
      <p class="clause-p">(5) The tenant is not entitled to make structural modifications to the rented rooms.</p>
    </div>
  </div>

  <div class="page-footer-num">5/7</div>
</div>

<!-- ================================= PAGE 6 ================================= -->
<div class="page" id="page-6">
  <div class="cols">
    <div class="col-de">
      <div class="clause-h" style="margin-top:0;">9. Betreten der Mietsache durch den Vermieter</div>
      <p class="clause-p">(1) Zur Überprüfung des Wohnungszustandes, der Feststellung der Notwendigkeit unaufschiebbarer Arbeiten, dem Ablesen der Messgeräte oder zum Zwecke des Verkaufs des Anwesens, ist der Vermieter berechtigt, die Mietsache zusammen mit Handwerkern bzw. Miet- oder Kaufinteressenten nach rechtzeitiger vorheriger Abstimmung mit dem Mieter, zu betreten.</p>
      <p class="clause-p">(2) Der Vermieter darf die Wohnung darüber hinaus betreten, wenn eine konkrete Gefahr für die Mietsache oder andere Mieter besteht.</p>
      <p class="clause-p">(3) Der Mieter hat dafür Sorge zu tragen, dass der Vermieter diese Rechte auch bei längerer Abwesenheit wahrnehmen kann.</p>

      <div class="clause-h">10. Rückgabe der Mietsache</div>
      <p class="clause-p">(1) Bei Ende des Mietvertrages hat der Mieter die vollständig möblierte Mietsache mit sämtlichen Schlüsseln – auch den vom Mieter zusätzlich beschafften – von privaten Gegenständen geräumt und besenrein zu übergeben.</p>
      <p class="clause-p">(2) Vom Mieter mitgebrachte Gegenstände sind zu entfernen.</p>

      <div class="clause-h">11. Erfüllungsort und Gerichtsstand</div>
      <p class="clause-p">Erfüllungsort und Gerichtsstand für alle sich aus diesem Vertrag ergebenden Verpflichtungen ist der Ort der Wohnung.</p>

      <div class="clause-h">12. Datenschutz</div>
      <p class="clause-p">Der Mieter hat davon Kenntnis genommen, dass die Daten, die zur ordnungsgemäßen Durchführung der vertraglichen Bestimmungen oder zur üblichen Betreuung des Mieters erforderlich sind, auf Datenträger gespeichert und nach den Bestimmungen der Datenschutzgrundverordnung (DSGVO) bearbeitet werden. Der Mieter erklärt sich mit der Einholung einer Auskunft über seine Zahlungsfähigkeit einverstanden.</p>

      <div class="clause-h">13. Individuelle Vereinbarungen</div>
      <p class="clause-p">{{INDIVIDUAL_AGREEMENTS_DE}}</p>
    </div>

    <div class="col-en">
      <div class="clause-h" style="margin-top:0;">9. Entitlement of the landlord to enter the rental property</div>
      <p class="clause-p">(1) In order to check the condition of the apartment, to determine the necessity of work that cannot be postponed, to read the counters, or for the purpose of selling the property, the landlord is entitled to enter the rental property together with craftspeople, prospective tenants or purchasers after timely prior consultation with the tenant.</p>
      <p class="clause-p">(2) Furthermore, the landlord may enter the apartment in case of concrete danger to the rental property or other tenants.</p>
      <p class="clause-p">(3) The tenant must ensure that the landlord can exercise these rights even in case of an extended absence.</p>

      <div class="clause-h">10. Return of the rental property</div>
      <p class="clause-p">(1) Upon termination of the rental agreement, the tenant must return the fully furnished rental property with all keys (including those additionally produced by the tenant), cleared of any private items and swept clean.</p>
      <p class="clause-p">(2) Objects brought by the tenant must be removed.</p>

      <div class="clause-h">11. Place of performance and jurisdiction</div>
      <p class="clause-p">The place of performance and the jurisdiction for all obligations arising from this contract is the location of the apartment.</p>

      <div class="clause-h">12. Data protection</div>
      <p class="clause-p">The tenant is aware that the data required for the proper execution of the contractual provisions or for the regular support of the tenant will be stored on data carriers and processed in accordance with the provisions of the General Data Protection Regulations (GDPR). The tenant agrees to obtain information about his/her solvency.</p>

      <div class="clause-h">13. Individual stipulations</div>
      <p class="clause-p">{{INDIVIDUAL_AGREEMENTS_EN}}</p>
    </div>
  </div>

  <div class="page-footer-num">6/7</div>
</div>

<!-- ================================= PAGE 7 ================================= -->
<div class="page" id="page-7">
  <div class="cols">
    <div class="col-de">
      <div class="clause-h" style="margin-top:0;">14. Salvatorische Klausel</div>
      <p class="clause-p">Soweit eine der Bestimmungen dieses Mietvertrages, gleich aus welchem Grund, unwirksam sein sollte, gelten die übrigen Bestimmungen unverändert fort. Die Vertragspartner vereinbaren jedoch bereits jetzt eine unwirksame Bestimmung durch eine gültige Bestimmung zu ersetzen, die dem gewollten Zweck in gesetzlich zulässiger Weise am nächsten kommt.</p>
    </div>

    <div class="col-en">
      <div class="clause-h" style="margin-top:0;">14. Severability clause</div>
      <p class="clause-p">Should one of the provisions of this rental agreement be invalid for any reason whatsoever, the remaining provisions shall continue to apply, unchanged. Nevertheless, the contracting parties hereby agree to replace an invalid provision with a valid provision which comes closest to the intended purpose in a lawful manner.</p>
    </div>
  </div>

  <table class="sign-table">
    <tr>
      <td>
        <div style="font-weight: bold; color: #000;">Kylasam Nallagoundanpalayam Arruchamy</div>
        <div style="color: #444; font-size: 8.8pt; margin: 3px 0 16px 0;">{{LANDLORD_SIGN_DATE}}</div>
        <div style="font-weight: bold; font-size: 9.5pt; color: #111;">Vom Vermieter akzeptiert</div>
        <div style="font-size: 8.8pt; color: #555;">Accepted by Landlord</div>
      </td>
      <td>
        <div style="font-weight: bold; color: #000;">{{TENANT_NAME}}</div>
        <div style="color: #444; font-size: 8.8pt; margin: 3px 0 16px 0;">{{TENANT_SIGN_DATE}}</div>
        <div style="font-weight: bold; font-size: 9.5pt; color: #111;">Vom Mieter akzeptiert</div>
        <div style="font-size: 8.8pt; color: #555;">Accepted by Tenant</div>
      </td>
    </tr>
  </table>

  <div class="page-footer-num">7/7</div>
</div>

</body>
</html>"""

def render_contract_html(data):
    tmpl = generate_html_template()
    for k, v in data.items():
        tmpl = tmpl.replace(f"{{{{{k}}}}}", str(v))
    return tmpl

def render_contract_docx(data, output_path):
    doc = Document()
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.6)
        s.bottom_margin = Inches(0.6)
        s.left_margin = Inches(0.65)
        s.right_margin = Inches(0.65)

    # PAGE 1
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_de = p_title.add_run("Mietvertrag\n")
    r_de.bold = True
    r_de.font.size = Pt(24)
    r_en = p_title.add_run("Rental Agreement")
    r_en.font.size = Pt(16)
    r_en.font.color.rgb = RGBColor(80, 80, 80)
    doc.add_paragraph()

    t_landlord = doc.add_table(rows=5, cols=2)
    t_landlord.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_landlord.columns[0].width = Inches(3.5)
    t_landlord.columns[1].width = Inches(3.5)
    
    t_landlord.cell(0, 0).paragraphs[0].add_run("Zwischen Vermieter:").bold = True
    t_landlord.cell(0, 1).paragraphs[0].add_run("The landlord:").bold = True
    t_landlord.cell(1, 0).paragraphs[0].add_run("Vorname, Nachname").font.size = Pt(8.5)
    t_landlord.cell(1, 1).paragraphs[0].add_run("Name, Surname").font.size = Pt(8.5)
    t_landlord.cell(2, 0).paragraphs[0].add_run("Kylasam NA").bold = True
    t_landlord.cell(2, 1).paragraphs[0].add_run("Kylasam NA").bold = True
    t_landlord.cell(3, 0).paragraphs[0].add_run("Anschrift").font.size = Pt(8.5)
    t_landlord.cell(3, 1).paragraphs[0].add_run("Address").font.size = Pt(8.5)
    t_landlord.cell(4, 0).paragraphs[0].add_run("Forckenbeckstraße 64,\n14199 Berlin, Deutschland")
    t_landlord.cell(4, 1).paragraphs[0].add_run("Forckenbeckstraße 64,\n14199 Berlin, Germany")
    doc.add_paragraph()

    t_tenant = doc.add_table(rows=5, cols=2)
    t_tenant.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_tenant.columns[0].width = Inches(3.5)
    t_tenant.columns[1].width = Inches(3.5)
    
    t_tenant.cell(0, 0).paragraphs[0].add_run("Zwischen Mieter:").bold = True
    t_tenant.cell(0, 1).paragraphs[0].add_run("The tenant:").bold = True
    t_tenant.cell(1, 0).paragraphs[0].add_run("Vorname, Nachname").font.size = Pt(8.5)
    t_tenant.cell(1, 1).paragraphs[0].add_run("Name, Surname").font.size = Pt(8.5)
    t_tenant.cell(2, 0).paragraphs[0].add_run(str(data["TENANT_NAME"])).bold = True
    t_tenant.cell(2, 1).paragraphs[0].add_run(str(data["TENANT_NAME"])).bold = True
    t_tenant.cell(3, 0).paragraphs[0].add_run("Anschrift").font.size = Pt(8.5)
    t_tenant.cell(3, 1).paragraphs[0].add_run("Address").font.size = Pt(8.5)
    t_tenant.cell(4, 0).paragraphs[0].add_run(str(data["TENANT_ADDRESS_DE"]).replace("<br>", "\n"))
    t_tenant.cell(4, 1).paragraphs[0].add_run(str(data["TENANT_ADDRESS_EN"]).replace("<br>", "\n"))

    doc.add_page_break()

    # PAGES 2-7 CLAUSES TABLE HELPER
    def add_dual_row(table, de_text, en_text, is_heading=False):
        row = table.add_row()
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(3.5)
        c1.width = Inches(3.5)
        p0 = c0.paragraphs[0]
        p1 = c1.paragraphs[0]
        r0 = p0.add_run(de_text.replace("<br>", "\n"))
        r1 = p1.add_run(en_text.replace("<br>", "\n"))
        if is_heading:
            r0.bold = True
            r1.bold = True
            r0.font.size = Pt(10)
            r1.font.size = Pt(10)
        else:
            r0.font.size = Pt(8.5)
            r1.font.size = Pt(8.5)

    # PAGE 2: Preamble, Clause 1, Clause 2
    t2 = doc.add_table(rows=0, cols=2)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_dual_row(t2, "Präambel", "Preamble", is_heading=True)
    add_dual_row(t2, 
        "Der Vermieter ist nicht zwingend Eigentümer des Mietobjektes, sondern hat den Mietgegenstand gegebenenfalls vom Eigentümer oder sonstigen Berechtigten zum Zwecke der Weitervermietung angemietet. Der Vermieter vermietet möblierten Wohnraum ausschließlich für begrenzte Zeit zum vorübergehenden Gebrauch.",
        "The landlord is not necessarily the owner of the rental property but has rented the property from the owner or another authorised party for the purpose of subletting it. The landlord rents a furnished living space exclusively for a limited period of time for temporary use."
    )
    add_dual_row(t2, "1. Mietsache", "1. Rental property", is_heading=True)
    add_dual_row(t2,
        f"(1) Der Vermieter vermietet dem Mieter zur Nutzung durch folgende Person(en)\n1. {data['TENANT_NAME']}\nzum vorübergehenden Gebrauch folgende Wohnung im Hause Lichtenrader Damm 80, 12305 Berlin, Deutschland , im {data['ROOM_FLOOR']} Geschoss bestehend aus 1 Zimmer(n).\nDie Wohnfläche beträgt ca. {data['ROOM_SIZE']} m². Diese Angabe dient wegen möglicher Messfehler nicht zur Festlegung des Mietgegenstandes. Der räumliche Umfang der gemieteten Sache ergibt sich vielmehr aus der Anzahl der vermieteten Räume.\n(2) Das Mietobjekt ist voll möbliert und komplett eingerichtet.\n(3) Bei Wohnungsübergabe werden dem Mieter folgende Schlüssel ausgehändigt:\nWie bei Übergabe vereinbart",
        f"(1) The landlord rents to the tenant for temporary use on behalf of the following person(s)\n1. {data['TENANT_NAME']}\nthe following apartment is located in Lichtenrader Damm 80, 12305 Berlin, Germany , on {data['ROOM_FLOOR']} floor consisting of 1 rooms.\nThe residential space consists of approximately {data['ROOM_SIZE']} m². Due to possible measurement errors, this information is not conducive to define the rental property. Instead, the size of the rental property is determined by the number of rented rooms.\n(2) The rental property is entirely furnished and completely equipped.\n(3) When the apartment is handed over, the following keys are provided to the tenant:\nAs agreed during key handover"
    )
    add_dual_row(t2, "2. Mietdauer", "2. Rental duration", is_heading=True)
    add_dual_row(t2,
        f"(1) Der Mietvertrag beginnt am {data['START_DATE_DE']} und endet am {data['END_DATE_DE']}.\n(2) Die Wohnung ist gemäß § 549 Abs. 2 Nr. 1 BGB lediglich zum vorübergehenden Gebrauch vermietet. Der Vermieter vermietet das Mietobjekt ausschließlich im Rahmen des vorübergehenden Gebrauchs. Auf Seiten des Mieters besteht ein vorübergehender Sonderbedarf, welcher durch die Anmietung der voll möblierten Mietsache im Rahmen des § 549 Abs. 2 Nr. 1 BGB gedeckt werden soll. Der Mieter bestätigt einen vorübergehenden Sonderbedarf {data['REASON_DE']}\nSofern der vorübergehende Wohnbedarf aufgrund geänderter Umstände über den vereinbarten Zeitraum hinaus fortbesteht, kommt eine Verlängerung des Mietverhältnisses in Betracht. Hierfür hat der Mieter dem Vermieter den Grund zum vorübergehenden Gebrauch spätestens einen vollen Monat vor Mietende mitzuteilen. Der Vermieter ist nicht verpflichtet, den neuen oder aktualisierten Grund für den Aufenthalt zu akzeptieren und das Mietverhältnis fortzusetzen.\n(3) Setzt der Mieter nach Ablauf der Mietzeit den Gebrauch der Mietsache fort, so verlängert sich das",
        f"(1) The rental contract starts on {data['START_DATE_EN']} and ends on {data['END_DATE_EN']}.\n(2) The apartment is rented for temporary use only, in accordance with section 549, paragraph 2(1) of the German Civil Code (Bürgerliches Gesetzbuch, BGB), the landlord rents the rental property exclusively for temporary use. The tenant has a special need for temporary accommodation, to be covered by renting the fully furnished rental property pursuant to section 549, paragraph 2(1) of the BGB. The tenant confirms a temporary special need {data['REASON_EN']}\nIf the temporary housing need continues beyond the agreed period due to changed circumstances, an extension of the rental period is possible. For this purpose, the tenant must inform the landlord of the reason for temporary use no later than one full month before the end of the rental period. The landlord is not obliged to accept the new or updated reason for residency and to resume the rental agreement.\n(3) If the tenant continues to use the rental property after the termination of the rental period, the rental relationship shall not be extended indefinitely. Section 545 of the BGB is excluded."
    )

    doc.add_page_break()

    # PAGE 3: Clause 3, Clause 4
    t3 = doc.add_table(rows=0, cols=2)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_dual_row(t3, "Mietverhältnis nicht auf unbestimmte Zeit. § 545 BGB wird ausgeschlossen.", "")
    add_dual_row(t3, "3. Miete", "3. Rent", is_heading=True)
    add_dual_row(t3,
        f"(1) Die Miete beträgt monatlich {data['RENT_AMOUNT']}€, in Worten: {data['RENT_WORDS_DE']} EURO.\nHierin enthalten ist\n• eine Nettokaltmiete inklusive Möblierungszuschlag\n• Nebenkosten inkl. Betriebs-, Heiz-, Warmwasser-, Strom- und Internetkosten\n(2) Der Mieter hat für eine sachgemäße, nicht übertriebene Nutzung (Fair Usage) der Nebenkosten zu sorgen. Hier ist besonders auf den sachgemäßen Verbrauch von Strom, Wasser und Heizung zu achten.\nDie Gesamtmiete beinhaltet eine Internet-Flatrate. Weitere Kosten, die durch die Nutzung der Internetverbindung entstehen (z.B. kostenpflichtige oder illegale Downloads etc.), sind vom Mieter der Wohnung selbst zu tragen und werden mit dem Vermieter abgerechnet. Beim Einzug stellt der Vermieter dem Mieter den Router und die Zugangsdaten für die WLAN-Verbindung zur Verfügung. Beide Parteien sind sich einig, dass der Mieter allein für die Nutzung der Internetverbindungen und die illegale Internetnutzung verantwortlich ist. Bei illegaler Nutzung kann der Vermieter die Internetverbindung ohne Vorankündigung sperren. In diesem Fall hat der Mieter keinen Anspruch auf Mietminderung.\nWeitere, unmittelbar mit der Wohnung zusammenhängende Verbrauchskosten, z. B. Telefonkosten, trägt der Mieter selbst.\n(3) Bei einem Mietverhältnis mit einer Dauer von unter 6 Monaten findet der aktuell gesetzlich gültige Umsatzsteuersatz Anwendung.\n(4) Der Mieter hat zusätzlich zur laufenden Miete eine Endreinigungspauschale in Höhe von {data['CLEANING_FEE']}€ zu leisten, welche bei Rückgabe der Wohnung bar zu übergeben oder zu überweisen ist oder nach Absprache mit dem Vermieter von der Kaution abgezogen wird.",
        f"(1) The rent amounts to €{data['RENT_AMOUNT']} per month, in words: {data['RENT_WORDS_EN']} EURO.\nThis includes\n• basic rent (ger:, Nettokaltmiete) which includes a furniture surcharge\n• Additional expenses lump sum incl. operating, utility, heating, electricity and internet cost\n(2) The tenant must ensure proper, non-exaggerated use (fair usage) of the utility costs in particular a proper consumption of electricity, water and heating is important.\nThe total rent includes an Internet flat rate. Further costs arising from the use of the Internet connection (e.g. fee-based or illegal downloads, etc.) are to be borne by the tenant of the apartment and will be settled with the landlord. When moving in, the landlord provides the tenant with the router and the login details to connect to the WLAN. Both parties agree that the tenant is solely responsible for the Internet connection and liable for the illegal use of the Internet. In the event of illegal use, the landlord can block the Internet connection without prior notice. In this case, the tenant cannot claim a reduction in rent.\nAdditional consumption costs directly related to the apartment (e.g. telephone costs) shall be borne by the tenant.\n(3) In the case of a rental agreement with a duration of less than 6 months, the currently valid statutory VAT rate shall apply.\n(4) In addition to the regular rent, the tenant must pay a fee for the final clean-up of €{data['CLEANING_FEE']}, which must be paid in cash or by bank transfer at the time of handover of the apartment, or otherwise deducted from the deposit upon consultation with the landlord."
    )
    add_dual_row(t3, "4. Mietzahlungen", "4. Payment of rent", is_heading=True)
    add_dual_row(t3,
        f"(1) Die erste Miete ist bei Vertragsunterzeichnung fällig und beläuft sich auf €{data['FIRST_RENT_AMOUNT']}. Leistet der Mieter die vereinbarte Miete nicht, besteht kein Anspruch auf Übergabe der Mietsache.\n(2) Im Folgenden ist die Miete monatlich im Voraus, spätestens bis zum 1. Werktag des jeweiligen Monats, kostenfrei an den Vermieter auf folgendes Konto zu leisten:",
        f"(1) The first rent is due after contract signing and amounts to {data['FIRST_RENT_AMOUNT']}€. If the tenant does not pay the agreed rent, there is no entitlement to occupy the rental property.\n(2) In the subsequent period, the rent is to be paid on a monthly basis in advance, at the latest by the first working day of the corresponding month, free of any charges to the landlord, to the following account:"
    )

    doc.add_page_break()

    # PAGE 4: Bank Details, Clause 5
    t4 = doc.add_table(rows=0, cols=2)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_dual_row(t4,
        f"Kontoinhaber: Kylasam Nallagoundanpalayam Arruchamy\nIBAN: DE87100110012624542768\nBIC: NTSBDEB1XXX\nVerwendungszweck: {data['TENANT_NAME']} \"Monat / Jahr\"\n\n(3) Für die Rechtzeitigkeit der Zahlung kommt es nicht auf die Absendung, sondern auf den Eingang des Geldes an.\n(4) Der Mieter hat für jede Mahnung/ Zahlungserinnerung wegen verspäteter Zahlung der Miete {data['LATE_FEE']} Euro für pauschalierte Mahnkosten zu zahlen, sofern der Mieter nicht nachweist, dass dem Vermieter geringere bzw. keine Kosten entstanden sind.",
        f"Account holder: Kylasam Nallagoundanpalayam Arruchamy\nIBAN: DE87100110012624542768\nBIC: NTSBDEB1XXX\nPayment reference: {data['TENANT_NAME']} \"Month / Year\"\n\n(3) Punctual payment does not depend on the time of transfer, but on the receipt of the money.\n(4) For each reminder/overdue payment notice attributable to the late payment of the rent, the tenant must pay a fee of €{data['LATE_FEE']} as a reminder cost, unless the tenant proves that the landlord has incurred lower or no expenses."
    )
    add_dual_row(t4, "5. Mietsicherheit/ Kaution", "5. Security deposit", is_heading=True)
    add_dual_row(t4,
        f"(1) Der Mieter leistet für die Erfüllung seiner Pflichten aus diesem Mietvertrag zu Beginn des Mietverhältnisses eine Sicherheit in Höhe von {data['DEPOSIT_AMOUNT']}€, an den Vermieter auf folgendes Konto:\n\nKontoinhaber: Kylasam Nallagoundanpalayam Arruchamy\nIBAN: DE25120400000011466000\nBIC: COBADEFFXXX\nVerwendungszweck: {data['TENANT_NAME']} \"Kaution\"\noder in bar bei Mietvertragsschluss.\n(2) Leistet der Mieter die vereinbarte Mietsicherheit nicht, besteht kein Anspruch auf Übergabe der Mietsache.\n(3) Nach Vertragsende und Rückgabe der Mietsache rechnet der Vermieter die Kaution gegenüber dem Mieter ab und zahlt, wenn sich die Wohnung in dem Zustand befindet, in dem sie sich am Anfang befand (ausgenommen Verschleiß) den errechneten Kautionsbetrag aus. Ein Zurückbehaltungsrecht des Vermieters besteht nicht, solange keine begründeten Gegenansprüche dargelegt werden (z.B. Schäden an Möbeln, nicht bezahlte Miete).",
        f"(1) For the fulfilment of his/ her obligations from this rental agreement, the tenant shall pay a security deposit of €{data['DEPOSIT_AMOUNT']}, to the landlord due upon the beginning of rental period, to the following bank account:\n\nAccount holder: Kylasam Nallagoundanpalayam Arruchamy\nIBAN: DE25120400000011466000\nBIC: COBADEFFXXX\nPayment reference: {data['TENANT_NAME']} \"Deposit\"\nor in cash upon conclusion of the rental contract.\n(2) If the tenant does not pay the agreed security deposit, there is no entitlement to occupy the rental property.\n(3) After termination of the contract and the return of the rental property, the landlord settles the security deposit with the tenant and pays back the deposit amount if the apartment is in the same condition as it was at the start of the contract (excluding wear and tear). The landlord cannot exercise a retention right as long as no justified counter-claims are presented (e.g. damage to furniture, unpaid rent)."
    )

    doc.add_page_break()

    # PAGE 5: Clause 6, Clause 7, Clause 8
    t5 = doc.add_table(rows=0, cols=2)
    t5.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_dual_row(t5, "6. Kündigung", "6. Termination", is_heading=True)
    add_dual_row(t5,
        "Der Mietvertrag wird für den in Ziffer 2 genannten Zeitraum fest abgeschlossen.\nDas Recht der Parteien zur außerordentlichen Kündigung des Mietverhältnisses bleibt hiervon unberührt.",
        "The rental contract is concluded for the period specified in Section 2.\nThe right of the parties to an exceptional termination of the rent remains unaffected."
    )
    add_dual_row(t5, "7. Nutzung der Mietsache", "7. Use of the rental property", is_heading=True)
    add_dual_row(t5,
        "(1) Eine weitere Untervermietung, sonstige Gebrauchsüberlassung an Dritte oder gewerbliche Nutzung ist dem Mieter ausdrücklich nicht gestattet.\n(2) Haustiere sind in der Wohnung nicht erlaubt\n(3) Rauchen ist in der Wohnung nicht erlaubt",
        "(1) A further subletting, other transfer of use to third parties or commercial use on behalf of the tenant is expressly prohibited.\n(2) Pets are not allowed in the apartment\n(3) Smoking in the apartment is not allowed"
    )
    add_dual_row(t5, "8. Instandhaltungspflicht/ Bauliche Veränderungen", "8. Obligation to undertake maintenance/ structural modifications", is_heading=True)
    add_dual_row(t5,
        "(1) Der Mieter ist verpflichtet, die Mietsache sowie die überlassenen Möbel, Einrichtungen und Anlagen schonend und pfleglich zu behandeln. Er hat für eine ordnungsgemäße Reinigung der Mietsache und für eine ausreichende Lüftung und Beheizung des ihm überlassenen Hauses zu sorgen.\n(2) Zeigt sich ein wesentlicher Mangel der Mietsache oder wird eine Vorkehrung zum Schutz der Mietsache oder des Anwesens gegen eine nicht vorhergesehene Gefahr erforderlich, so hat der Mieter dies dem Vermieter unverzüglich mitzuteilen und ihm Gelegenheit zu geben, innerhalb angemessener Frist für Abhilfe zu sorgen. Der Mieter hat keinen Anspruch auf Ersatz seiner Aufwendungen, wenn er nicht dem Vermieter vorher diese Gelegenheit gegeben hat, es sei denn, es bestand Gefahr im Verzug.\n(3) Den Mieter trifft die Beweislast für den Umstand, dass er einen Schaden nicht zu vertreten hat, wenn die Schadensursache in einem Bereich der Mietsache liegt, der nur ihm als Mietpartei zugänglich ist.\n(4) Hat der Mieter einen Schaden zu vertreten, ist er verpflichtet, diesen sofort zu beseitigen. Das bedeutet eine angemessene Schadensersatzzahlung oder der Austausch des beschädigten Inventars.\n(5) Der Mieter ist nicht berechtigt, an den gemieteten Räumen bauliche Veränderungen vorzunehmen.",
        "(1) The tenant is obliged to treat the rental property, as well as the furniture, furnishings and equipment with respect and care. The tenant shall ensure proper cleaning and sufficient ventilation and heating of the rental property.\n(2) If a substantial defect of the rental property becomes apparent or if a precautionary measure is necessary to protect the rental property or the premises against an unforeseen hazard, the tenant shall immediately inform the landlord and grant the opportunity to remedy the situation within a reasonable period of time. The tenant is not entitled to compensation for his/her expenses if no notice was given to the landlord concerning this opportunity, unless an imminent danger was present.\n(3) The burden of proof, stating that he/she is not responsible for damage, is borne by the tenant if the cause of the damage is located in an area of the rental property that is solely accessible to him/her as a rental party.\n(4) If the tenant is responsible for damage, he/she is obliged to remedy it immediately. This means an appropriate compensation or the replacement of the damaged inventory.\n(5) The tenant is not entitled to make structural modifications to the rented rooms."
    )

    doc.add_page_break()

    # PAGE 6: Clause 9, 10, 11, 12, 13
    t6 = doc.add_table(rows=0, cols=2)
    t6.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_dual_row(t6, "9. Betreten der Mietsache durch den Vermieter", "9. Entitlement of the landlord to enter the rental property", is_heading=True)
    add_dual_row(t6,
        "(1) Zur Überprüfung des Wohnungszustandes, der Feststellung der Notwendigkeit unaufschiebbarer Arbeiten, dem Ablesen der Messgeräte oder zum Zwecke des Verkaufs des Anwesens, ist der Vermieter berechtigt, die Mietsache zusammen mit Handwerkern bzw. Miet- oder Kaufinteressenten nach rechtzeitiger vorheriger Abstimmung mit dem Mieter, zu betreten.\n(2) Der Vermieter darf die Wohnung darüber hinaus betreten, wenn eine konkrete Gefahr für die Mietsache oder andere Mieter besteht.\n(3) Der Mieter hat dafür Sorge zu tragen, dass der Vermieter diese Rechte auch bei längerer Abwesenheit wahrnehmen kann.",
        "(1) In order to check the condition of the apartment, to determine the necessity of work that cannot be postponed, to read the counters, or for the purpose of selling the property, the landlord is entitled to enter the rental property together with craftspeople, prospective tenants or purchasers after timely prior consultation with the tenant.\n(2) Furthermore, the landlord may enter the apartment in case of concrete danger to the rental property or other tenants.\n(3) The tenant must ensure that the landlord can exercise these rights even in case of an extended absence."
    )
    add_dual_row(t6, "10. Rückgabe der Mietsache", "10. Return of the rental property", is_heading=True)
    add_dual_row(t6,
        "(1) Bei Ende des Mietvertrages hat der Mieter die vollständig möblierte Mietsache mit sämtlichen Schlüsseln – auch den vom Mieter zusätzlich beschafften – von privaten Gegenständen geräumt und besenrein zu übergeben.\n(2) Vom Mieter mitgebrachte Gegenstände sind zu entfernen.",
        "(1) Upon termination of the rental agreement, the tenant must return the fully furnished rental property with all keys (including those additionally produced by the tenant), cleared of any private items and swept clean.\n(2) Objects brought by the tenant must be removed."
    )
    add_dual_row(t6, "11. Erfüllungsort und Gerichtsstand", "11. Place of performance and jurisdiction", is_heading=True)
    add_dual_row(t6,
        "Erfüllungsort und Gerichtsstand für alle sich aus diesem Vertrag ergebenden Verpflichtungen ist der Ort der Wohnung.",
        "The place of performance and the jurisdiction for all obligations arising from this contract is the location of the apartment."
    )
    add_dual_row(t6, "12. Datenschutz", "12. Data protection", is_heading=True)
    add_dual_row(t6,
        "Der Mieter hat davon Kenntnis genommen, dass die Daten, die zur ordnungsgemäßen Durchführung der vertraglichen Bestimmungen oder zur üblichen Betreuung des Mieters erforderlich sind, auf Datenträger gespeichert und nach den Bestimmungen der Datenschutzgrundverordnung (DSGVO) bearbeitet werden. Der Mieter erklärt sich mit der Einholung einer Auskunft über seine Zahlungsfähigkeit einverstanden.",
        "The tenant is aware that the data required for the proper execution of the contractual provisions or for the regular support of the tenant will be stored on data carriers and processed in accordance with the provisions of the General Data Protection Regulations (GDPR). The tenant agrees to obtain information about his/her solvency."
    )
    add_dual_row(t6, "13. Individuelle Vereinbarungen", "13. Individual stipulations", is_heading=True)
    add_dual_row(t6,
        str(data["INDIVIDUAL_AGREEMENTS_DE"]).replace("<br>", "\n"),
        str(data["INDIVIDUAL_AGREEMENTS_EN"]).replace("<br>", "\n")
    )

    doc.add_page_break()

    # PAGE 7: Clause 14, Signatures
    t7 = doc.add_table(rows=0, cols=2)
    t7.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_dual_row(t7, "14. Salvatorische Klausel", "14. Severability clause", is_heading=True)
    add_dual_row(t7,
        "Soweit eine der Bestimmungen dieses Mietvertrages, gleich aus welchem Grund, unwirksam sein sollte, gelten die übrigen Bestimmungen unverändert fort. Die Vertragspartner vereinbaren jedoch bereits jetzt eine unwirksame Bestimmung durch eine gültige Bestimmung zu ersetzen, die dem gewollten Zweck in gesetzlich zulässiger Weise am nächsten kommt.",
        "Should one of the provisions of this rental agreement be invalid for any reason whatsoever, the remaining provisions shall continue to apply, unchanged. Nevertheless, the contracting parties hereby agree to replace an invalid provision with a valid provision which comes closest to the intended purpose in a lawful manner."
    )
    doc.add_paragraph()

    t_sign = doc.add_table(rows=4, cols=2)
    t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_sign.columns[0].width = Inches(3.5)
    t_sign.columns[1].width = Inches(3.5)
    
    t_sign.cell(0, 0).paragraphs[0].add_run("Kylasam Nallagoundanpalayam Arruchamy").bold = True
    t_sign.cell(0, 1).paragraphs[0].add_run(str(data["TENANT_NAME"])).bold = True
    t_sign.cell(1, 0).paragraphs[0].add_run(str(data["LANDLORD_SIGN_DATE"])).font.size = Pt(8.5)
    t_sign.cell(1, 1).paragraphs[0].add_run(str(data["TENANT_SIGN_DATE"])).font.size = Pt(8.5)
    t_sign.cell(2, 0).paragraphs[0].add_run("Vom Vermieter akzeptiert").bold = True
    t_sign.cell(2, 1).paragraphs[0].add_run("Vom Mieter akzeptiert").bold = True
    t_sign.cell(3, 0).paragraphs[0].add_run("Accepted by Landlord").font.size = Pt(8.5)
    t_sign.cell(3, 1).paragraphs[0].add_run("Accepted by Tenant").font.size = Pt(8.5)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)

def convert_html_to_pdf(html_path, pdf_path):
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
    chrome_cmd = [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "--headless",
        "--disable-gpu",
        "--print-to-pdf-no-header",
        f"--print-to-pdf={pdf_path}",
        html_path
    ]
    res = subprocess.run(chrome_cmd, capture_output=True, text=True)
    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
        return pdf_path
    else:
        raise RuntimeError(f"Chrome PDF generation failed: {res.stderr}")

def get_gcloud_token():
    try:
        return subprocess.check_output(["gcloud", "auth", "print-access-token"]).decode().strip()
    except Exception as e:
        print(f"⚠ Could not get gcloud token: {e}")
        return None

def upload_or_update_drive_file(token, folder_id, file_path, doc_name, mime_type):
    with open(file_path, "rb") as f:
        file_bytes = f.read()

    # Check if file with same name exists in this folder
    query = f"'{folder_id}' in parents and name = '{doc_name}' and trashed = false"
    list_url = f"https://www.googleapis.com/drive/v3/files?q={urllib.parse.quote(query)}&fields=files(id,name)"
    req = urllib.request.Request(list_url, headers={"Authorization": f"Bearer {token}"})
    existing_id = None
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode())
            if data.get("files"):
                existing_id = data["files"][0]["id"]
    except Exception as e:
        print(f"  ⚠ Query error: {e}")

    boundary = "===============HOMELYO_DRIVE_BOUNDARY=="
    metadata = {
        "name": doc_name,
        "mimeType": mime_type
    }

    if existing_id:
        # Update existing file
        upload_url = f"https://www.googleapis.com/upload/drive/v3/files/{existing_id}?uploadType=multipart"
        method = "PATCH"
    else:
        # Create new file in folder
        metadata["parents"] = [folder_id]
        upload_url = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart"
        method = "POST"

    body = (
        f"--{boundary}\r\n"
        f"Content-Type: application/json; charset=UTF-8\r\n\r\n"
        f"{json.dumps(metadata)}\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: {mime_type}\r\n\r\n"
    ).encode("utf-8") + file_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")

    req = urllib.request.Request(
        upload_url,
        data=body,
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
        method=method
    )
    with urllib.request.urlopen(req) as resp:
        res = json.loads(resp.read().decode())
        action = "Updated" if existing_id else "Created"
        print(f"  ✓ {action} Google Drive file: '{doc_name}' (ID: {res.get('id')}) in folder {folder_id}")
        return res.get("id")

def main():
    print("Building official 7-page Homelyo Rental Contract (exact layout of SIGNED_CONTRACT_67da88a52ff5e11a1f9d8ffa)...")

    # Sample master data (Bhargavi reference values for master template)
    sample_data = render_contract_data(
        tenant_name="Bhargavi Mangesh Ayare",
        tenant_email="bhargavi.ayare@example.com",
        tenant_phone="+491512345678",
        room_code="KE2",
        start_date="2025-03-25",
        end_date="2025-09-30",
        rent_amount=639.00,
        deposit_amount=1200.00,
        tenant_address_de="Type 4 building no 3 flat no 4 esis hospital quarters , Vashi sector-5 sagar vihar road, 400703 Navi Mumbai, Indien",
        tenant_address_en="Type 4 building no 3 flat no 4 esis hospital quarters , Vashi sector-5 sagar vihar road, 400703 Navi Mumbai, India",
        reason_de="Aufgrund eines Studiensemesters außerhalb seines eigentlichen Studienortes.",
        reason_en="Due to a semester of study outside of his/her normal place of study.",
        landlord_sign_date="20 Mar 2025, 10:39 GMT",
        tenant_sign_date="21 Mar 2025, 12:36 GMT"
    )

    # 1. Generate Master HTML Template
    raw_html = generate_html_template()
    with open(HTML_TEMPLATE_PATH, "w", encoding="utf-8") as f:
        f.write(raw_html)
    print(f"✓ Master HTML template saved: {HTML_TEMPLATE_PATH}")

    # 2. Generate Master DOCX
    render_contract_docx(sample_data, DOCX_TEMPLATE_PATH)
    print(f"✓ Master DOCX template saved: {DOCX_TEMPLATE_PATH}")

    # 3. Generate Sample Master PDF
    rendered_html = render_contract_html(sample_data)
    temp_html = "/tmp/master_contract_preview.html"
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(rendered_html)
    convert_html_to_pdf(temp_html, FINAL_PDF_PATH)
    print(f"✓ Master PDF generated: {FINAL_PDF_PATH}")

    # 4. Check PDF page count
    try:
        import fitz
        doc = fitz.open(FINAL_PDF_PATH)
        print(f"✓ Verification: Master PDF page count is {len(doc)} pages (Target: 7)")
    except Exception as e:
        print(f"ℹ PyMuPDF check: {e}")

    # 5. Sync to Google Drive
    token = get_gcloud_token()
    if token:
        print(f"\n📂 Syncing Templates to Google Drive folder '{TEMPLATES_FOLDER_ID}'...")
        upload_or_update_drive_file(
            token, TEMPLATES_FOLDER_ID, DOCX_TEMPLATE_PATH, "template_contract.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        upload_or_update_drive_file(
            token, TEMPLATES_FOLDER_ID, HTML_TEMPLATE_PATH, "contract_template.html",
            "text/html"
        )
        upload_or_update_drive_file(
            token, TEMPLATES_FOLDER_ID, FINAL_PDF_PATH, "Rental_Contract_Master_Final.pdf",
            "application/pdf"
        )

        print(f"\n📂 Syncing Sample Contract to Generated Contracts folder '{GENERATED_CONTRACTS_FOLDER_ID}'...")
        upload_or_update_drive_file(
            token, GENERATED_CONTRACTS_FOLDER_ID, FINAL_PDF_PATH,
            "Rental_Contract_Bhargavi_Mangesh_Ayare_KE2_20250325_20250930.pdf",
            "application/pdf"
        )
        upload_or_update_drive_file(
            token, GENERATED_CONTRACTS_FOLDER_ID, DOCX_TEMPLATE_PATH,
            "Rental_Contract_Bhargavi_Mangesh_Ayare_KE2_20250325_20250930.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        print("✓ All templates and generated contracts synchronized with Google Drive!")
    else:
        print("⚠ GCloud token unavailable; skipped remote Google Drive sync.")

if __name__ == "__main__":
    main()
