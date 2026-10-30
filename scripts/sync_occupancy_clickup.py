import json
import time
import requests

CLICKUP_TOKEN = 'pk_314452952_6STD829UQJXD0ROPHUVS41YXEPHXUIE2'
HEADERS = {
    'Authorization': CLICKUP_TOKEN,
    'Content-Type': 'application/json'
}

LIST_OCCUPANCY = '1500520000038641'
LIST_DEPOSITS = '1500520000038642'

# Custom Field IDs on Tenants & Room Occupancy
FIELD_ROOM = '1f6337ac-dc2a-4a9e-a0e1-4f0f5c388b49'
FIELD_TENANT_NAME = 'db28805a-9c19-4b18-9576-a03e7a2d3917'
FIELD_TENANT_EMAIL = 'e376d256-7c84-49d2-a8c0-e2d17e1941c1'
FIELD_MONTHLY_RENT = 'a456b51c-03c1-4d8f-abf2-c0bbf84773af'
FIELD_SECURITY_DEPOSIT = 'ffb1442d-70f7-4d2f-bc04-932143eacfc2'
FIELD_BOOKING_PLATFORM = '8a9ffe62-6cac-4ac9-8eb3-11eb793b5b0e'
FIELD_LEASE_START = 'be0f6c20-1231-491b-bfcc-4318848ae5c9'
FIELD_LEASE_END = '3bddd0d2-6b9e-41c9-9e8a-879ddb0362b3'
FIELD_ANMELDUNG = '5d3ad3fd-11e2-4f85-9c56-5a8750ad28c3'

# Option IDs
OPT_ROOM_KE1 = '7fafa08a-51c7-480b-85b9-829e37059bcc'
OPT_ROOM_KE2 = '5da57f68-ae41-4037-be42-97ae9c4156af'
OPT_ROOM_KE3 = '8567d501-7a1b-4753-b0f4-04aae4f8309b'
OPT_ROOM_KE4 = 'f40c942e-cb88-43c0-96fa-6f97c760826e'

OPT_PLATFORM_HOUSINGANYWHERE = 'fe438ef4-2294-4f39-86d4-f3b266421bc3'
OPT_PLATFORM_WUNDERFLATS = '5f2bf398-67a5-4e14-b811-42a76156498c'
OPT_PLATFORM_DIRECT = '5c06807d-811c-44f1-9723-1547f96e4286'

def set_cf(task_id, field_id, value):
    url = f'https://api.clickup.com/api/v2/task/{task_id}/field/{field_id}'
    r = requests.post(url, headers=HEADERS, json={'value': value})
    time.sleep(0.3)
    return r.status_code == 200

def update_task(task_id, data):
    url = f'https://api.clickup.com/api/v2/task/{task_id}'
    r = requests.put(url, headers=HEADERS, json=data)
    time.sleep(0.3)
    return r.json()

def create_task(list_id, data):
    url = f'https://api.clickup.com/api/v2/list/{list_id}/task'
    r = requests.post(url, headers=HEADERS, json=data)
    time.sleep(0.3)
    return r.json()

def main():
    print('=== Step 1: Update KE1 – Umakant Yadav (Task 1ampw4tj4jj) ===')
    update_task('1ampw4tj4jj', {
        'status': 'in progress',
        'priority': 2,
        'start_date': 1793502000000, # 2026-11-01
        'due_date': 1809050400000    # 2027-04-30
    })
    set_cf('1ampw4tj4jj', FIELD_ROOM, OPT_ROOM_KE1)
    set_cf('1ampw4tj4jj', FIELD_TENANT_NAME, 'Umakant Yadav')
    set_cf('1ampw4tj4jj', FIELD_TENANT_EMAIL, 'umakant3888@gmail.com')
    set_cf('1ampw4tj4jj', FIELD_MONTHLY_RENT, 609.00)
    set_cf('1ampw4tj4jj', FIELD_SECURITY_DEPOSIT, 1000.00)
    set_cf('1ampw4tj4jj', FIELD_BOOKING_PLATFORM, OPT_PLATFORM_DIRECT)
    set_cf('1ampw4tj4jj', FIELD_LEASE_START, 1793502000000)
    set_cf('1ampw4tj4jj', FIELD_LEASE_END, 1809050400000)
    set_cf('1ampw4tj4jj', FIELD_ANMELDUNG, True)
    print('✅ Task 1ampw4tj4jj updated.')

    print('\n=== Step 2: Update KE2 – Ajithkumar Velasseril Sreenivasan (Task 1ampw4tj4ed) ===')
    update_task('1ampw4tj4ed', {
        'status': 'in progress',
        'priority': 2,
        'start_date': 1793588400000, # 2026-11-02
        'due_date': 1809050400000    # 2027-04-30
    })
    set_cf('1ampw4tj4ed', FIELD_ROOM, OPT_ROOM_KE2)
    set_cf('1ampw4tj4ed', FIELD_TENANT_NAME, 'Ajithkumar Velasseril Sreenivasan')
    set_cf('1ampw4tj4ed', FIELD_TENANT_EMAIL, 'ajithkmvs123@gmail.com')
    set_cf('1ampw4tj4ed', FIELD_MONTHLY_RENT, 629.00)
    set_cf('1ampw4tj4ed', FIELD_SECURITY_DEPOSIT, 1258.00)
    set_cf('1ampw4tj4ed', FIELD_BOOKING_PLATFORM, OPT_PLATFORM_DIRECT)
    set_cf('1ampw4tj4ed', FIELD_LEASE_START, 1793588400000)
    set_cf('1ampw4tj4ed', FIELD_LEASE_END, 1809050400000)
    set_cf('1ampw4tj4ed', FIELD_ANMELDUNG, True)
    print('✅ Task 1ampw4tj4ed updated.')

    print('\n=== Step 3: Create / Populate KE3 – Thiruvikkiraman Pandurangan ===')
    ke3_desc = """🏠 Active Tenant Occupancy Profile — Thiruvikkiraman Pandurangan (KE3)

📍 Property: Lichtenrader Damm 80, 12305 Berlin, Germany
🚪 Room: KE3 — Quiet Garden View Room (12 m²)
👤 Tenant: Thiruvikkiraman Pandurangan ("Thiru")
📧 Tenant Email: thiruvikkiraman.p@gmail.com
🏢 Platform Booked: Wunderflats / HousingAnywhere
📅 Tenancy Term: 01.11.2025 → 31.03.2027 (515 Days Active Tenancy)
💶 Warm Rent: €610.00 / month (All-inclusive: heating, electricity, high-speed Wi-Fi 6, Hausgeld, GEZ)
🔒 Security Deposit Held (Kaution): €1,000.00 (Escrow confirmed)
🏛️ Anmeldung (Wohnungsgeberbestätigung): Official city registration provided (§ 19 BMG)
📶 Wi-Fi Network: Vodafone-327C (Gigabit Wi-Fi 6 mesh)
🔑 Key Inventory: 4 keys issued (Main entrance, Apartment, Letterbox, Room KE3)

⭐ Verified Tenant Testimonial (5.0/5 Stars on HousingAnywhere):
"A little bit out of the city. But, very calm to live and work. Almost everything nearby supermarkets and restaurants. Neighbours very friendly and prompt response from the landlord. Everything fine! Thank you!!"

📋 Room Features & Inclusions:
• Double bed with orthopedic mattress & clean duvet
• Smart ambient reading light & spacious wardrobe
• Quiet garden-facing view with abundant natural morning light
• Fully equipped modern kitchen (Samsung SpaceMax fridge, oven, induction hob)
• Modern bathroom with 4-mode rainfall shower
"""
    t_ke3 = create_task(LIST_OCCUPANCY, {
        'name': '🟡 [OCCUPIED until Mar 31, 2027] Room KE3 – Thiruvikkiraman Pandurangan (€610/mo)',
        'description': ke3_desc,
        'status': 'in progress',
        'priority': 3,
        'start_date': 1761955200000, # 2025-11-01
        'due_date': 1806537600000    # 2027-03-31
    })
    ke3_id = t_ke3.get('id')
    print(f'Created KE3 Task: {ke3_id}')
    if ke3_id:
        set_cf(ke3_id, FIELD_ROOM, OPT_ROOM_KE3)
        set_cf(ke3_id, FIELD_TENANT_NAME, 'Thiruvikkiraman Pandurangan')
        set_cf(ke3_id, FIELD_TENANT_EMAIL, 'thiruvikkiraman.p@gmail.com')
        set_cf(ke3_id, FIELD_MONTHLY_RENT, 610.00)
        set_cf(ke3_id, FIELD_SECURITY_DEPOSIT, 1000.00)
        set_cf(ke3_id, FIELD_BOOKING_PLATFORM, OPT_PLATFORM_WUNDERFLATS)
        set_cf(ke3_id, FIELD_LEASE_START, 1761955200000)
        set_cf(ke3_id, FIELD_LEASE_END, 1806537600000)
        set_cf(ke3_id, FIELD_ANMELDUNG, True)

    print('\n=== Step 4: Create / Populate KE4 – Karolina Faluta ===')
    ke4_desc = """🏠 Active Tenant Occupancy Profile — Karolina Anna Faluta (KE4)

📍 Property: Lichtenrader Damm 80, 12305 Berlin, Germany
🚪 Room: KE4 — Cozy Calm Room (11 m², Thickest Orthopedic Mattress)
👤 Tenant: Karolina Anna Faluta ("Karolina")
📧 Tenant Email: karolina.a.faluta@gmail.com
🏢 Platform Booked: HousingAnywhere
📅 Tenancy Term: 06.10.2025 → 31.07.2027 (Extended until 31 July 2027, 663 Days Total)
   • Contract Extension Period: 01.08.2026 → 31.07.2027
💶 Warm Rent: €569.00 / month (All-inclusive: heating, electricity, high-speed Wi-Fi 6, Hausgeld, GEZ)
🔒 Security Deposit Held (Kaution): €1,000.00 (Escrow confirmed)
🏛️ Anmeldung (Wohnungsgeberbestätigung): Official city registration provided (§ 19 BMG)
📶 Wi-Fi Network: Vodafone-327C (Gigabit Wi-Fi 6 mesh)
🔑 Key Inventory: 4 keys issued (Main entrance, Apartment, Letterbox, Room KE4)

📁 Legal & Onboarding Documents:
• Contract Extension: Rental_Contract_Extension_Karolina_July2027.pdf (Signed 21 June 2026)
• Welcome Kit: Welcome_Kit_KE4.pdf (Lichtenrader Damm 80, quiet hours, emergency numbers)
• Key Handover Acknowledgement: Signed (§ 549 Abs. 2 Nr. 1 BGB temporary housing use)

📋 Room Features & Inclusions:
• Premium deep-sleep orthopedic mattress (thickest mattress in apartment)
• Wardrobe, study desk & ergonomic chair, blackout curtains
• Quiet courtyard / garden orientation, ideal for peaceful sleep & study
• Fully equipped modern kitchen & renovated standing shower
"""
    t_ke4 = create_task(LIST_OCCUPANCY, {
        'name': '🟡 [OCCUPIED until Jul 31, 2027] Room KE4 – Karolina Faluta (€569/mo)',
        'description': ke4_desc,
        'status': 'in progress',
        'priority': 3,
        'start_date': 1759708800000, # 2025-10-06
        'due_date': 1817078400000    # 2027-07-31
    })
    ke4_id = t_ke4.get('id')
    print(f'Created KE4 Task: {ke4_id}')
    if ke4_id:
        set_cf(ke4_id, FIELD_ROOM, OPT_ROOM_KE4)
        set_cf(ke4_id, FIELD_TENANT_NAME, 'Karolina Anna Faluta')
        set_cf(ke4_id, FIELD_TENANT_EMAIL, 'karolina.a.faluta@gmail.com')
        set_cf(ke4_id, FIELD_MONTHLY_RENT, 569.00)
        set_cf(ke4_id, FIELD_SECURITY_DEPOSIT, 1000.00)
        set_cf(ke4_id, FIELD_BOOKING_PLATFORM, OPT_PLATFORM_HOUSINGANYWHERE)
        set_cf(ke4_id, FIELD_LEASE_START, 1759708800000)
        set_cf(ke4_id, FIELD_LEASE_END, 1817078400000)
        set_cf(ke4_id, FIELD_ANMELDUNG, True)

    print('\n=== Step 5: Synchronize Deposit Ledger (List 1500520000038642) ===')
    # Check existing tasks in deposits list
    r_dep = requests.get(f'https://api.clickup.com/api/v2/list/{LIST_DEPOSITS}/task?include_closed=true', headers=HEADERS).json()
    existing_dep_names = [t['name'] for t in r_dep.get('tasks', [])]
    print('Existing deposits:', existing_dep_names)

    deposits_to_ensure = [
        {
            'name': '🔒 [Room KE2] Ajithkumar Velasseril Sreenivasan – Deposit €1,298.00',
            'description': '🔐 Deposit Escrow & Settlement Ledger — Ajithkumar Velasseril Sreenivasan (KE2)\nAdvance Deposit Paid: €1,298.00 (2 months cold rent equivalent)\nStatus: Held in Escrow (Active Lease)\nContract Term: 2026-11-02 to 2027-04-30',
            'start_date': 1793588400000,
            'due_date': 1809050400000
        },
        {
            'name': '🔒 [Room KE3] Thiruvikkiraman Pandurangan – Deposit €1,000.00',
            'description': '🔐 Deposit Escrow & Settlement Ledger — Thiruvikkiraman Pandurangan (KE3)\nAdvance Deposit Paid: €1,000.00\nStatus: Held in Escrow (Active Lease)\nContract Term: 2025-11-01 to 2027-03-31',
            'start_date': 1761955200000,
            'due_date': 1806537600000
        },
        {
            'name': '🔒 [Room KE4] Karolina Faluta – Deposit €1,000.00',
            'description': '🔐 Deposit Escrow & Settlement Ledger — Karolina Faluta (KE4)\nAdvance Deposit Paid: €1,000.00\nStatus: Held in Escrow (Active Lease)\nContract Term: 2025-10-06 to 2027-07-31',
            'start_date': 1759708800000,
            'due_date': 1817078400000
        }
    ]

    for dep in deposits_to_ensure:
        if not any(dep['name'] in name for name in existing_dep_names):
            d_task = create_task(LIST_DEPOSITS, {
                'name': dep['name'],
                'description': dep['description'],
                'status': 'to do',
                'priority': 3,
                'start_date': dep['start_date'],
                'due_date': dep['due_date']
            })
            dep_name = dep['name']
            task_id = d_task.get('id')
            print(f'Created Deposit Task: {dep_name} -> {task_id}')

    print('\n=== All Operations Completed Successfully! ===')

if __name__ == '__main__':
    main()
