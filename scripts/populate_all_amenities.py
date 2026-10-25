import urllib.request
import json
import time

TOKEN = "pk_308598557_9CP65EANIWJ0XHF9NUOFGMFUDCAMPPM5"
LIST_ID = "1500350000002184" # Apartment Inventory & Assets
HEADERS = {"Authorization": TOKEN, "Content-Type": "application/json"}

amenities = [
    # 1. Kitchen Major Appliances
    ("Samsung Double-Door SpaceMax™ Refrigerator", "Common Kitchen", "Brand New", 799, "Samsung SpaceMax No Frost, Twin Cooling Plus, digital inverter, quiet 35 dB operation."),
    ("Digital Convection Microwave Oven", "Common Kitchen", "Brand New", 180, "Brand-new digital microwave with multi-mode defrost, auto-cook, and reheat presets."),
    ("Ceramic Induction Hob Stove & Convection Oven Combo", "Common Kitchen", "Brand New", 650, "Touch-control 4-zone induction hob and multi-function electric convection oven with baking trays."),
    ("Small Kitchen Trio: Kettle, Toaster & Sandwich Panini Maker", "Common Kitchen", "Brand New", 140, "Stainless steel 1.7L rapid-boil kettle, 2-slice automatic toaster, non-stick sandwich press."),

    # 2. Kitchen Cookware & Prep
    ("Non-Stick & Stainless Steel Pots & Pans Cookware Set", "Common Kitchen", "Brand New", 180, "Heavy-gauge non-stick frying pans, 4 stainless cooking pots with heat-resistant tempered glass lids."),
    ("Oven Baking Trays & Roasting Dishes", "Common Kitchen", "Brand New", 65, "Non-stick oven baking trays, roasting pans, and heat-resistant ovenware."),
    ("Cooking Utensils Suite (Ladle, Spatula, Whisk, Tongs)", "Common Kitchen", "Brand New", 45, "Heat-resistant silicone and stainless cooking utensils, pasta fork, slotted spoon."),
    ("Professional Chef & Paring Knife Block Set", "Common Kitchen", "Brand New", 85, "High-carbon stainless steel chef knife, bread knife, utility knife, paring knife, wooden block."),
    ("Bamboo & Heavy-Duty Wooden Cutting Boards", "Common Kitchen", "Brand New", 35, "Hygienic multi-board cutting set for vegetables, bread, and meal preparation."),
    ("Kitchen Gadgets (Can Opener, Bottle Opener, Peeler, Corkscrew)", "Common Kitchen", "Brand New", 30, "Heavy-duty ergonomic can opener, bottle opener, vegetable peeler, and corkscrew."),
    ("Colander, Mesh Strainer & Salad Serving Suite", "Common Kitchen", "Brand New", 35, "Stainless steel colander, fine mesh strainer, and salad serving spoons."),

    # 3. Dining, Glassware & Tableware
    ("Porcelain Dinner & Salad Plates Service (12-Piece)", "Common Kitchen", "Brand New", 90, "Durable white porcelain dinner plates and side/salad plates for all 4 housemates and guests."),
    ("Soup & Cereal Porcelain Bowls Set (8-Piece)", "Common Kitchen", "Brand New", 50, "Deep porcelain bowls for cereal, soups, and shared meals."),
    ("Drinking Glasses & Highball Tumblers (12-Piece)", "Common Kitchen", "Brand New", 40, "Clear glass water and juice tumblers, dishwasher safe."),
    ("Ceramic Coffee & Tea Mugs Set", "Common Kitchen", "Brand New", 35, "Comfortable handle ceramic mugs for morning tea and coffee."),
    ("Polished Stainless Steel Cutlery & Flatware (24-Piece)", "Common Kitchen", "Brand New", 75, "Forks, table knives, soup spoons, and dessert spoons."),
    ("Kitchen Towels, Drying Cloths & Silicone Oven Mitts", "Common Kitchen", "Brand New", 25, "100% cotton absorbent dish towels and silicone heat-proof oven gloves."),
    ("Countertop Dish Drying Rack & Sink Caddy", "Common Kitchen", "Brand New", 40, "Rust-proof stainless dish drying rack with removable drainage tray and utensil caddy."),

    # 4. Bathroom Spa Renovation
    ("Standing Shower System with 4 Spray Modes (Rain, Massage, Mist, Jet)", "Shared Bathroom", "Brand New", 1450, "Luxury standing shower with 4 adjustable flow modes: Refreshing Rain, Deep Massage, Soft Mist, Jet."),
    ("Bathroom Wall & Floor Designer Ceramic Tiles", "Shared Bathroom", "Brand New", 2200, "Fully renovated bathroom with brand-new premium wall tiles and non-slip designer floor tiles."),
    ("Illuminated LED Mirror Cabinet with Internal Storage", "Shared Bathroom", "Brand New", 320, "Anti-fog illuminated mirror cabinet, internal shelving, integrated shaver/toothbrush charging socket."),
    ("Contemporary Floating Vanity Unit with Soft-Close Drawers", "Shared Bathroom", "Brand New", 550, "Modern vanity with integrated ceramic washbasin, high-neck chrome mixer tap, soft-close storage."),
    ("Upgraded Plumbing, Pressure Booster & Pipe Network", "Shared Bathroom", "Brand New", 1200, "Brand-new plumbing lines delivering strong water pressure, reliability, and silent operation."),
    ("High-Efficiency Bathroom Extractor & Ventilation System", "Shared Bathroom", "Brand New", 240, "Automatic humidity-sensing ventilation fan keeping the bathroom fresh and mold-free."),
    ("Heated Towel Radiator Rail, Hooks & Anti-Slip Bath Mats", "Shared Bathroom", "Brand New", 180, "Wall-mounted heated towel dryer, stainless robe hooks, and absorbent memory foam bath mats."),

    # 5. Private Rooms
    ("Room KE1: King/Queen Bed & Orthopedic Mattress (21 m²)", "Room KE1", "Brand New", 650, "Master bedroom bed frame with deep orthopedic spring mattress and fresh bedding package."),
    ("Room KE1: Large Executive Study Desk & Ergonomic Chair", "Room KE1", "Brand New", 320, "Spacious workstation for remote work and studying, adjustable ergonomic office chair."),
    ("Room KE1: Spacious Double Wardrobe with Full-Length Mirror", "Room KE1", "Brand New", 450, "Large wardrobe with coat hanging rail, organizers, and full-length dressing mirror."),
    ("Room KE1: Dimmable Ambient Mood Lighting & Blackout Curtains", "Room KE1", "Brand New", 160, "Floor-standing ambient lamp, desk lamp, and thermal blackout drapes for quality sleep."),
    ("Room KE2: Private Balcony Outdoor Bistro Table & Chairs", "Room KE2", "Brand New", 220, "Weatherproof balcony seating set overlooking the peaceful private garden courtyard."),
    ("Room KE2: Double Bed, Orthopedic Mattress & Work Desk", "Room KE2", "Brand New", 1200, "Balcony suite bedroom setup: double bed, work desk, chair, wardrobe, and balcony door blinds."),
    ("Room KE3: Double Bed, Wardrobe & Smart Music-Sync Mood Lighting", "Room KE3", "Brand New", 1150, "Garden-view bedroom with double bed, wardrobe, desk, and smart music-synchronized ambient lighting."),
    ("Room KE4: Ultra-Thick Orthopedic Reset Mattress & Bedroom Suite", "Room KE4", "Brand New", 1100, "Peaceful courtyard room with the thickest orthopedic reset mattress, desk, and wardrobe."),

    # 6. Common Storage & Infrastructure
    ("Dedicated Common Luggage & Suitcase Storage Room", "Luggage Storage / Corridor", "Brand New", 350, "Separate common storage room equipped with heavy-duty metal racks for travel trolleys and bulky items."),
    ("Rowenta Silence Force High-Power Vacuum Cleaner", "Luggage Storage / Corridor", "Brand New", 190, "Multi-surface quiet vacuum cleaner with HEPA filtration system for apartment hygiene."),
    ("Steam Iron & Folding Adjustable Ironing Board", "Luggage Storage / Corridor", "Brand New", 85, "Non-drip steam iron and heat-reflective folding ironing board."),
    ("Apartment Floor Care & Cleaning Essentials Kit", "Luggage Storage / Corridor", "Brand New", 60, "Vileda microfiber mop, bucket, broom, dustpan, and eco-friendly starter cleaning kit."),
    ("AVM FRITZ!Box 6660 Cable Gigabit Wi-Fi 6 Mesh Router", "Luggage Storage / Corridor", "Brand New", 240, "Ultra-fast Gigabit Wi-Fi 6 mesh network delivering up to 1,000 Mbps seamless coverage across all rooms."),
    ("Optical Smoke Alarms & Emergency First Aid Kit", "All Rooms & Hallways", "Brand New", 150, "DIN EN 14604 certified smoke alarms in every room, fire blanket, and comprehensive first aid kit.")
]

# Fetch existing tasks to avoid duplicates
existing_tasks = {}
url_get = f"https://api.clickup.com/api/v2/list/{LIST_ID}/task?include_closed=true"
res = urllib.request.urlopen(urllib.request.Request(url_get, headers=HEADERS))
for t in json.loads(res.read().decode()).get("tasks", []):
    existing_tasks[t["name"]] = t["id"]

print(f"Existing items in Amenities: {len(existing_tasks)}")

for title, location, condition, value, desc in amenities:
    time.sleep(0.5)
    if title in existing_tasks:
        print(f"  ⏭️ Exists: {title}")
        continue

    full_desc = f"""📍 Location: {location}
🏷️ Condition: {condition}
💰 Estimated Replacement Value: €{value:.2f}
📦 Part of Lichtenrader Damm 80 Co-Living Move-In Package

Details:
{desc}"""

    payload = {
        "name": f"[{location}] {title}",
        "description": full_desc,
        "priority": 3
    }

    url_post = f"https://api.clickup.com/api/v2/list/{LIST_ID}/task"
    req = urllib.request.Request(url_post, data=json.dumps(payload).encode("utf-8"), headers=HEADERS)
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode())
            print(f"  ✅ Added Amenity: [{data.get('id')}] [{location}] {title}")
    except urllib.error.HTTPError as e:
        print(f"  ❌ Error {title}: {e.code} {e.read().decode()[:150]}")

print("=== Amenities Population Finished! ===")
