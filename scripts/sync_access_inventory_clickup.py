#!/usr/bin/env python3
"""
Synchronizes the Table-Style Tenant Access & Resource Inventory across all
tenant profiles and the Master Inventory task in ClickUp.
"""

import json
import urllib.request

TOKEN = "pk_314452952_6STD829UQJXD0ROPHUVS41YXEPHXUIE2"
HEADERS = {"Authorization": TOKEN, "Content-Type": "application/json"}
LIST_OCCUPANCY = "1500520000038641"

def update_task_md(task_id, md_text):
    url = f"https://api.clickup.com/api/v2/task/{task_id}"
    req = urllib.request.Request(
        url,
        data=json.dumps({"markdown_description": md_text}).encode("utf-8"),
        headers=HEADERS,
        method="PUT"
    )
    with urllib.request.urlopen(req) as resp:
        print(f"✓ Updated ClickUp task {task_id} (HTTP {resp.status})")

# 1. Room KE1 - Umakant Yadav
ke1_md = """# 🏠 Tenant Onboarding & Access Profile — Umakant Yadav (KE1)

**Property**: Lichtenrader Damm 80, 12305 Berlin · **Management**: `support@homelyo.eu`  
**Room**: KE1 (Master Bedroom, 21 m²) · **Platform**: HousingAnywhere  
**Lease Term**: `01.11.2026` → `30.04.2027` (6 Months) · **Rent**: €609.00/mo · **Deposit**: €1,000.00  
**Tenant Email**: `umakant3888@gmail.com` · **Phone**: +49 15222 319536

---

## 🔐 Table-Style Access & Resource Inventory (Room KE1)

| Resource Category | Resource / Asset Name | Direct Link / Identifier | Authorized Access | Permission / Status |
| :--- | :--- | :--- | :--- | :--- |
| **Mobile Hub** | Room KE1 Resident Portal | [linktr.ee/lichtenraderdamm_80_ke_1](https://linktr.ee/lichtenraderdamm_80_ke_1) | Umakant Yadav | ✅ Live (Executive Theme) |
| **Support & Repairs** | Service & Maintenance Request Form | [homelyo.eu/assistant](https://www.homelyo.eu/assistant/) | Umakant Yadav | ✅ Live on Linktree & Kit |
| **Cloud Vault** | Personal PDF Folder (`My_docs`) | [Open `My_docs` Vault](https://drive.google.com/drive/folders/1x6WqAM3ap1xi0EGKNWRcrafycA0y-wiU) | `umakant3888@gmail.com` | 🔒 Restricted (Share Oct 31) |
| **Contract** | Signed Mietvertrag (PDF) | [View Contract PDF](https://drive.google.com/file/d/1_VZQHkXwq0rAk27GiND1mD5fd8ZN1gDU/view) | Admin + Umakant Yadav | ✅ Verified in `My_docs` |
| **Anmeldung** | Wohnungsgeberbestätigung (§ 19 BMG) | [View Anmeldung PDF](https://drive.google.com/file/d/1R-fiY2Wb6ZI7rKb0HSxKVmi4gfSzUGcs/view) | Admin + Umakant Yadav | ✅ Verified in `My_docs` |
| **Key Receipt** | Keys Acknowledgement Protocol | [View Keys PDF](https://drive.google.com/file/d/1U-Ya3CQ6jCHY_w8mjd0keoTmADkk1ba0/view) | Admin + Umakant Yadav | ✅ Verified in `My_docs` |
| **Welcome Kit** | Executive 2-Page Welcome Kit | [View Welcome Kit PDF](https://drive.google.com/file/d/10aVMgMfbv8a9Yjv687dy2U3D-SDODorU/view) | Admin + Umakant Yadav | ✅ QR Encodes `My_docs` |
| **House Rules** | Hausordnung & Quiet Hours | [View House Rules PDF](https://drive.google.com/file/d/1jzoiMO3qPPJh3Od6PLv3Iyhj3yFCm1iE/view) | All Residents | ✅ Verified in `My_docs` |
| **Handover** | Apartment Handover Protocol | [View Protocol PDF](https://drive.google.com/file/d/1xhQxYA_uW7sylx4r8xlnIl-1az8a6TxC/view) | Admin + Umakant Yadav | ✅ Verified in `My_docs` |
| **Inventory** | Room & Common Inventory List | [View Inventory PDF](https://drive.google.com/file/d/1fSxlHt9tPe9n0U1PZTunRdoaTBgAbetE/view) | Admin + Umakant Yadav | ✅ Verified in `My_docs` |
| **Transit Map** | Berlin S+U-Bahn Map (S2) | [View Transit Map PDF](https://drive.google.com/file/d/1ZlahgvnfgOLQarVMLjvQgRB9X-QM8hKs/view) | All Residents | ✅ Verified in `My_docs` |
| **Physical Keys** | 4 Keys (Main, Apt, Room KE1, Mailbox) | Physical Handover | Umakant Yadav | ⏳ Handover on `01.11.2026` |
"""

# 2. Room KE2 - Vacant / Available
ke2_md = """# 🟢 Room KE2 — Private Balcony Suite (24 m²)

**Property**: Lichtenrader Damm 80, 12305 Berlin · **Management**: `support@homelyo.eu`  
**Status**: Available from `01.11.2026` · **Warm Rent**: €629.00/mo · **Deposit**: €1,258.00

---

## 🔐 Table-Style Access & Resource Inventory (Room KE2)

| Resource Category | Resource / Asset Name | Direct Link / Identifier | Authorized Access | Permission / Status |
| :--- | :--- | :--- | :--- | :--- |
| **Mobile Hub** | Room KE2 Resident Portal | [linktr.ee/lichtenraderdamm_80_ke_2](https://linktr.ee/lichtenraderdamm_80_ke_2) | Future KE2 Tenant | ✅ Live (Executive Theme) |
| **Support & Repairs** | Service & Maintenance Request Form | [homelyo.eu/assistant](https://www.homelyo.eu/assistant/) | Future KE2 Tenant | ✅ Live on Linktree |
| **Building Master** | Lichtenrader Damm 80 Master Hub | [linktr.ee/lichtenrder_damm_80](https://linktr.ee/lichtenrder_damm_80) | All Residents | ✅ Active |
| **Cloud Vault** | Personal PDF Folder (`My_docs`) | Auto-provisioned upon booking | Admin Only | 🟢 Ready for Onboarding |
| **Physical Keys** | 4 Keys (Main, Apt, Room KE2, Mailbox) | Safe Storage | Homelyo Management | 🟢 In Safe Storage |
"""

# 3. Room KE3 - Thiruvikkiraman Pandurangan
ke3_md = """# 🏠 Tenant Onboarding & Access Profile — Thiruvikkiraman Pandurangan (KE3)

**Property**: Lichtenrader Damm 80, 12305 Berlin · **Management**: `support@homelyo.eu`  
**Room**: KE3 (Quiet Garden View, 12 m²) · **Platform**: Wunderflats  
**Lease Term**: `01.11.2026` → `31.03.2027` (5 Months) · **Rent**: €610.00/mo · **Deposit**: €1,220.00

---

## 🔐 Table-Style Access & Resource Inventory (Room KE3)

| Resource Category | Resource / Asset Name | Direct Link / Identifier | Authorized Access | Permission / Status |
| :--- | :--- | :--- | :--- | :--- |
| **Mobile Hub** | Room KE3 Resident Portal | [linktr.ee/lichtenraderdamm_80_ke_3](https://linktr.ee/lichtenraderdamm_80_ke_3) | Thiruvikkiraman Pandurangan | ✅ Live (Executive Theme) |
| **Support & Repairs** | Service & Maintenance Request Form | [homelyo.eu/assistant](https://www.homelyo.eu/assistant/) | Thiruvikkiraman Pandurangan | ✅ Live on Linktree |
| **Building Master** | Lichtenrader Damm 80 Master Hub | [linktr.ee/lichtenrder_damm_80](https://linktr.ee/lichtenrder_damm_80) | All Residents | ✅ Active |
| **House Rules** | Hausordnung & Quiet Hours | [View House Rules PDF](https://drive.google.com/file/d/15rkiK8grB2nYpFZfQs7RTakmcBFM0A_k/view) | All Residents | ✅ Linked on Portal |
| **Transit Map** | Berlin S+U-Bahn Transit Map | [View Transit Map PDF](https://drive.google.com/file/d/1z_ao-a8JTagdBPMaz-sw380FCr59HNov/view) | All Residents | ✅ Linked on Portal |
| **Physical Keys** | 4 Keys (Main, Apt, Room KE3, Mailbox) | Physical Handover Protocol | Thiruvikkiraman Pandurangan | ✅ Assigned for Tenancy |
"""

# 4. Room KE4 - Karolina Anna Faluta
ke4_md = """# 🏠 Tenant Onboarding & Access Profile — Karolina Anna Faluta (KE4)

**Property**: Lichtenrader Damm 80, 12305 Berlin · **Management**: `support@homelyo.eu`  
**Room**: KE4 (Cozy Calm Room, 11 m²) · **Platform**: HousingAnywhere  
**Lease Term**: `01.11.2026` → `31.07.2027` (9 Months) · **Rent**: €569.00/mo · **Deposit**: €1,138.00

---

## 🔐 Table-Style Access & Resource Inventory (Room KE4)

| Resource Category | Resource / Asset Name | Direct Link / Identifier | Authorized Access | Permission / Status |
| :--- | :--- | :--- | :--- | :--- |
| **Mobile Hub** | Room KE4 Resident Portal | [linktr.ee/lichtenraderdamm_80_ke_4](https://linktr.ee/lichtenraderdamm_80_ke_4) | Karolina Anna Faluta | ✅ Live (Executive Theme) |
| **Support & Repairs** | Service & Maintenance Request Form | [homelyo.eu/assistant](https://www.homelyo.eu/assistant/) | Karolina Anna Faluta | ✅ Live on Linktree |
| **Building Master** | Lichtenrader Damm 80 Master Hub | [linktr.ee/lichtenrder_damm_80](https://linktr.ee/lichtenrder_damm_80) | All Residents | ✅ Active |
| **House Rules** | Hausordnung & Quiet Hours | [View House Rules PDF](https://drive.google.com/file/d/15rkiK8grB2nYpFZfQs7RTakmcBFM0A_k/view) | All Residents | ✅ Linked on Portal |
| **Physical Keys** | 4 Keys (Main, Apt, Room KE4, Mailbox) | Physical Handover Protocol | Karolina Anna Faluta | ✅ Assigned for Tenancy |
"""

if __name__ == "__main__":
    update_task_md("1ampw4tj4jj", ke1_md)
    update_task_md("1ampw4tj4ed", ke2_md)
    update_task_md("1ampw4tjx3h", ke3_md)
    update_task_md("1ampw4tjx3j", ke4_md)
    update_task_md("1ampw4tj4j0", ke1_md)
    print("🎉 All ClickUp tenant access resource tables synchronized!")
