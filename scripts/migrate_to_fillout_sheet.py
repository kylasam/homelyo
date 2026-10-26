#!/usr/bin/env python3
"""
Migrate tenant records to target Google Sheet:
https://docs.google.com/spreadsheets/d/1gcQ4nv-lu8j3Kt3mNoFx11W7jkU0Umu1Ee2C2wM6Nvw/edit

Aligns sheet layout with Fillout form (gq1XFEqEWJus), removes older test rows,
and populates the 2 active tenant bookings (HousingAnywhere KE1 and Ajithkumar KE2).
"""

import sys
import json
import subprocess
import urllib.request
import urllib.error
import datetime

TARGET_SHEET_ID = "1gcQ4nv-lu8j3Kt3mNoFx11W7jkU0Umu1Ee2C2wM6Nvw"

FILLOUT_HEADERS = [
    "Submission ID",
    "Submission Date",
    "Full Name",
    "Email",
    "Phone Number",
    "Current Address",
    "Preferred Room",
    "Move-In Date (Preferred)",
    "Booking Period",
    "Length of Stay",
    "Number of Occupants",
    "Additional Requests or Question",
    "How did you hear about us?"
]

LIVE_ROWS = [
    [
        "sub_ha_ke1_20261101",
        "2026-10-26 12:00:00",
        "HousingAnywhere Tenant",
        "housinganywhere-booking-ke1@homelyo.eu",
        "+49 176 00000000",
        "HousingAnywhere Booking, Berlin",
        "KE1",
        "2026-11-01",
        "2026-11-01 - 2027-04-30",
        "6",
        "1",
        "Instant booking confirmed via HousingAnywhere platform. Folder: https://drive.google.com/drive/u/0/folders/1aOrC8nEXP7WCgldjcAIWCwiFLb1AFXRe",
        "Housing Anywhere"
    ],
    [
        "sub_ajith_ke2_20261102",
        "2026-10-26 14:30:00",
        "Ajithkumar Velasseril Sreenivasan",
        "ajithkmvs123@gmail.com",
        "+91 8606817191",
        "FU Berlin, Department of Physics, Arnimallee 14, 14195 Berlin, Germany",
        "KE2",
        "2026-11-02",
        "2026-11-02 - 2027-04-30",
        "6",
        "1",
        "Postdoc FU Berlin. Contract & Wohnungsgeberbestätigung needed. Folder: https://drive.google.com/drive/u/0/folders/1WMUDX6oqFx_7HbJSJGUes4m6Rwb_C86z",
        "Direct / Homelyo Portal"
    ]
]

def get_token():
    try:
        return subprocess.check_output(["gcloud", "auth", "print-access-token"], text=True).strip()
    except Exception as e:
        print(f"Error getting gcloud token: {e}")
        return None

def main():
    token = get_token()
    if not token:
        print("❌ Could not obtain gcloud access token.")
        sys.exit(1)

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # 1. Fetch spreadsheet metadata to get the first sheet's title
    meta_url = f"https://sheets.googleapis.com/v4/spreadsheets/{TARGET_SHEET_ID}"
    req = urllib.request.Request(meta_url, headers=headers)
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        print(f"❌ HTTP Error {e.code}: {e.reason}")
        print(e.read().decode())
        if e.code == 403:
            print("\n🚨 PERMISSION DENIED: The Google Sheet 1gcQ4nv-lu8j3Kt3mNoFx11W7jkU0Umu1Ee2C2wM6Nvw is not yet shared with kylashpriya@gmail.com.")
            print("👉 Please open https://docs.google.com/spreadsheets/d/1gcQ4nv-lu8j3Kt3mNoFx11W7jkU0Umu1Ee2C2wM6Nvw/edit")
            print("   Click 'Share' -> add kylashpriya@gmail.com as 'Editor' (or set 'Anyone with the link can edit').")
        sys.exit(1)

    sheet_title = data["sheets"][0]["properties"]["title"]
    print(f"✓ Found target spreadsheet: '{data.get('properties', {}).get('title')}' (first tab: '{sheet_title}')")

    # 2. Clear entire sheet data
    clear_url = f"https://sheets.googleapis.com/v4/spreadsheets/{TARGET_SHEET_ID}/values/{urllib.parse.quote(sheet_title)}!A1:Z100:clear"
    req = urllib.request.Request(clear_url, data=b"{}", headers=headers, method="POST")
    with urllib.request.urlopen(req) as resp:
        print(f"✓ Cleared older rows in '{sheet_title}'.")

    # 3. Write Fillout-aligned layout and 2 active records
    all_rows = [FILLOUT_HEADERS] + LIVE_ROWS
    update_url = f"https://sheets.googleapis.com/v4/spreadsheets/{TARGET_SHEET_ID}/values/{urllib.parse.quote(sheet_title)}!A1?valueInputOption=USER_ENTERED"
    payload = {"values": all_rows}
    req = urllib.request.Request(update_url, data=json.dumps(payload).encode("utf-8"), headers=headers, method="PUT")
    with urllib.request.urlopen(req) as resp:
        res = json.loads(resp.read().decode())
        print(f"✓ Successfully wrote {res.get('updatedRows', len(all_rows))} rows with Fillout header layout!")

    print("\n✅ Migration complete! Sheet URL: https://docs.google.com/spreadsheets/d/" + TARGET_SHEET_ID)

if __name__ == "__main__":
    main()
