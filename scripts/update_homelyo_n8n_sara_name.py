#!/usr/bin/env python3
import subprocess
import urllib.request
import json

def main():
    n8n_key = subprocess.check_output(
        ['gcloud', 'secrets', 'versions', 'access', 'latest', '--secret=N8N_API_KEY', '--project=n8n-homelyo'],
        text=True
    ).strip()

    base_url = 'https://n8n-fggh.onrender.com/api/v1'
    headers = {'X-N8N-API-KEY': n8n_key, 'Content-Type': 'application/json'}

    req = urllib.request.Request(f'{base_url}/workflows/db8UYqenoWzZV6qf', headers=headers)
    with urllib.request.urlopen(req) as resp:
        wf = json.loads(resp.read().decode())

    format_code = """const webhookData = $node["Webhook"].json.body || {};
const headers = $node["Webhook"].json.headers || {};
const geminiReply = $node["Gemini Flash API"].json.candidates?.[0]?.content?.parts?.[0]?.text || "No reply generated";
const userMessage = webhookData.message || "(empty message)";
const isLead = geminiReply.toLowerCase().includes("booking form") || geminiReply.toLowerCase().includes("booking request") || geminiReply.toLowerCase().includes("https://forms.fillout.com");

const sessionId = webhookData.session_id || 'web_' + (headers["cf-ray"] ? headers["cf-ray"].slice(0, 8) : Math.random().toString(36).substring(2, 8));
const turn = webhookData.turn ? `Turn #${webhookData.turn}` : 'Direct';
const country = headers["cf-ipcountry"] || headers["x-country"] || "DE";
const city = headers["cf-ipcity"] || "Berlin";

let intent = "general";
const msgLower = userMessage.toLowerCase();
if (msgLower.includes("anmeldung") || msgLower.includes("registration")) intent = "anmeldung_guarantee";
else if (msgLower.includes("book") || msgLower.includes("reserve") || msgLower.includes("apply") || msgLower.includes("vacan")) intent = "booking_request";
else if (msgLower.includes("rent") || msgLower.includes("price") || msgLower.includes("deposit") || msgLower.includes("kaution") || msgLower.includes("cost")) intent = "pricing_deposit";
else if (msgLower.includes("room") || msgLower.includes("available") || msgLower.includes("vacancy") || msgLower.includes("balcony")) intent = "room_availability";
else if (msgLower.includes("viewing") || msgLower.includes("visit") || msgLower.includes("tour")) intent = "virtual_viewing";
else if (msgLower.includes("pet") || msgLower.includes("dog") || msgLower.includes("cat")) intent = "pet_policy";
else if (msgLower.includes("train") || msgLower.includes("bus") || msgLower.includes("station") || msgLower.includes("supermarket")) intent = "location_transit";
else if (msgLower.includes("washer") || msgLower.includes("wifi") || msgLower.includes("internet") || msgLower.includes("kitchen")) intent = "amenities";

const leadBadge = isLead ? "\\n🔥 **HIGH-INTENT BOOKING LEAD**" : "";

const chatMarkdown = `🤖 **SARA CONCIERGE • HOMELYO LIVE DISPATCH**
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${leadBadge}
👤 **Visitor:** \\`${sessionId}\\` (${city}, ${country}) • ${turn}

💬 **Visitor Message:**
> *"${userMessage.replace(/\\n/g, ' ')}"*

🤖 **Sara AI Concierge Replied:**
${geminiReply}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📍 *Topic / Intent: \\`${intent}\\` • Homelyo Website AI*`;

return [{
  json: {
    user_message: userMessage,
    sara_reply: geminiReply,
    is_lead: isLead,
    session_id: sessionId,
    turn: turn,
    intent: intent,
    city: city,
    country: country,
    homelyo_chat_content: chatMarkdown
  }
}];"""

    process_content_code = """const items = $input.all();
let text = '';

for (const item of items) {
  if (item.json.content) {
    text += item.json.content + '\\n';
  }
}

// Section P: Verified Tenant Reviews & HousingAnywhere facts
text += `\\n============================================================\\nSECTION P: VERIFIED TENANT REVIEWS & HOUSINGANYWHERE VERIFICATION ⭐\\n============================================================\\nOfficial Credibility & Legitimacy Facts:\\n- Homelyo is 100% verified, legal, and legitimate.\\n- Address: Lichtenrader Damm 80, 12305 Berlin-Lichtenrade, Germany.\\n- All tenants receive official, legally binding German rental agreements and the official Wohnungsgeberbestätigung for Berlin city registration (Anmeldung).\\n- Rating: Straight 5.0/5 across Landlord, Communication, Move-in Process, Listing Accuracy, Cleanliness, Amenities, Furniture, and Location.\\n\\nActual Verified Tenant Reviews on HousingAnywhere:\\n1. Review from Naeem Ullah Khan (Stayed: Nov 2025 – Apr 2026, 5.0/5 Stars):\\n   - Listing Accuracy (5.0/5): "It was EXACTLY how it looked in the pictures. I was amazed. I had horrible experiences with properties from here. This was a pleasant change."\\n   - Move-in Process (5.0/5): "Went above and beyond to accommodate my early move in. Very pleased with the experience."\\n   - Landlord & Experience (5.0/5): "Met him in person on the day of move in. He was very nice and amazing person. Both for me and my girlfriend who moved into the same house with me in a different room."\\n   - Communication (5.0/5): "Always replied quickly."\\n   - Overall: "It was amazing. Both me and my girlfriend love the property, and our roommates and landlord both are very nice people so far. So we are happy."\\n\\n2. Review from Thiruvikkiraman Pandurangan (Stayed: Apr 2025 – Oct 2025, 5.0/5 Stars):\\n   - Ratings: Landlord 5.0/5, Communication 5.0/5, Move-in 5.0/5, Listing accuracy 5.0/5, Cleanliness 5.0/5, Amenities 5.0/5, Furniture 5.0/5, Location 5.0/5.\\n   - Comment: "A little bit out of the city. But, very calm to live and work. Almost everything nearby supermarkets and restaurants. Neighbours very friendly and prompt response from the landlord. Everything fine! Thank you!!"\\n\\nTenant Reviews Links:\\n- HousingAnywhere Tenant Reviews: https://housinganywhere.com/my/tenant-reviews\\n- Website Reviews & Testimonials: https://www.homelyo.eu/#reviews\\n\\nBooking Options & Price Transparency:\\n- Option 1: Book directly with Homelyo (via https://www.homelyo.eu/book or the in-chat booking request form).\\n- Option 2: Book via HousingAnywhere (our verified partner platform).\\n- The Key Difference:\\n  * On HousingAnywhere: The tenant must pay an extra platform/service fee charged by HousingAnywhere.\\n  * Directly with Homelyo: There are NO extra platform fees. The tenant pays ONLY the first month's rent and the advance refundable security deposit (Kaution).\\n  * Both booking options provide the exact same verified room, official rental contract, Anmeldung registration, and check-in support.\\n`;

// Section Q: Current Live Room Availability and Pricing
text += `\\n============================================================\\nSECTION Q: CURRENT LIVE ROOM AVAILABILITY & PRICING (OCTOBER 2026 UPDATE)\\n============================================================\\n- Room KE1: Master Bedroom (~21 m²), €639/month all-inclusive. CURRENT STATUS: Occupied & Booked from Nov 1, 2026 until April 30, 2027. EARLIEST NEXT AVAILABLE MOVE-IN: May 1, 2027.\\n- Room KE2: Premium Room with Private Balcony (~24 m²), ~~€649~~ €629 (November offer: €20 off)/month all-inclusive. CURRENT STATUS: Booked through Oct 31, 2026. EARLIEST NEXT AVAILABLE MOVE-IN: November 1, 2026 (Ready for Immediate Move-In!).\\n- Room KE3: Medium Bedroom (~16 m²), €619/month all-inclusive. CURRENT STATUS: Booked through March 31, 2027. EARLIEST NEXT AVAILABLE MOVE-IN: April 1, 2027.\\n- Room KE4: Cozy Value Bedroom (~14 m²), €569/month all-inclusive. CURRENT STATUS: Booked through July 31, 2027. EARLIEST NEXT AVAILABLE MOVE-IN: August 1, 2027.\\n`;

const userMessage = $node["Webhook"].json.body.message || '';
const customerData = $node["Webhook"].json.body.customer_data || null;

return [{
  json: {
    knowledge_base: text,
    user_message: userMessage,
    customer_data: customerData
  }
}];"""

    for n in wf['nodes']:
        if n['name'] == 'Format Chat Audit Log':
            n['parameters']['jsCode'] = format_code
        if n['name'] == 'Process Content':
            n['parameters']['jsCode'] = process_content_code

    update_payload = {
        'name': wf['name'],
        'nodes': wf['nodes'],
        'connections': wf['connections'],
        'settings': wf.get('settings', {})
    }

    put_req = urllib.request.Request(
        f'{base_url}/workflows/db8UYqenoWzZV6qf',
        data=json.dumps(update_payload).encode('utf-8'),
        headers=headers,
        method='PUT'
    )
    with urllib.request.urlopen(put_req) as resp:
        res = json.loads(resp.read().decode())
        print(f"Workflow db8UYqenoWzZV6qf updated successfully on Render: Active={res.get('active')}")

if __name__ == '__main__':
    main()
