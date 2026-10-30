#!/usr/bin/env python3
"""
Deploy updated Homelyo Chatbot n8n workflow with:
1. Clear Room Availability rules: Room KE1 is booked until May 1, 2027; ONLY Room KE2 is available for move-in.
2. Smart thread nesting in ClickUp Chat channel 1ampw4t88mv-92 using bot token pk_314452952_...
3. Preserves all verified tenant reviews (Naeem, Thiruvikkiraman) and Welcome Kit guidance.
"""
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
const turnNumber = parseInt(webhookData.turn || '1', 10);
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

const clickupToken = "pk_314452952_6STD829UQJXD0ROPHUVS41YXEPHXUIE2";
let parentMessageId = null;

// If this is Turn 2+, check if a parent thread exists for this session in ClickUp
if (turnNumber > 1) {
  try {
    const listRes = await this.helpers.httpRequest({
      method: 'GET',
      url: 'https://api.clickup.com/api/v3/workspaces/1500520000004763/chat/channels/1ampw4t88mv-92/messages?limit=25',
      headers: {
        'Authorization': clickupToken,
        'Content-Type': 'application/json',
        'accept': 'application/json'
      },
      json: true
    });
    const msgs = (listRes && listRes.data) ? listRes.data : [];
    for (const m of msgs) {
      if (m.content && m.content.includes(sessionId)) {
        parentMessageId = m.id;
        break;
      }
    }
  } catch (err) {
    // Fallback to top-level if error
    parentMessageId = null;
  }
}

let dispatchUrl = '';
let chatMarkdown = '';
const leadBadge = isLead ? "\\n🔥 **HIGH-INTENT BOOKING LEAD**" : "";

if (parentMessageId) {
  // Post as a THREAD REPLY under the existing visitor thread
  dispatchUrl = `https://api.clickup.com/api/v3/workspaces/1500520000004763/chat/messages/${parentMessageId}/replies`;
  chatMarkdown = `💬 **Visitor (Turn #${turnNumber}):**
> *"${userMessage.replace(/\\n/g, ' ')}"*

🤖 **Sara AI Concierge Replied:**
${geminiReply}

📍 *Intent: \\`${intent}\\`*`;
} else {
  // Post as a NEW TOP-LEVEL THREAD starter
  dispatchUrl = 'https://api.clickup.com/api/v3/workspaces/1500520000004763/chat/channels/1ampw4t88mv-92/messages';
  chatMarkdown = `🧵 **[NEW VISITOR CHAT THREAD]**
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${leadBadge}
👤 **Visitor:** \\`${sessionId}\\` (${city}, ${country}) • Turn #1

💬 **Visitor Message:**
> *"${userMessage.replace(/\\n/g, ' ')}"*

🤖 **Sara AI Concierge Replied:**
${geminiReply}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📍 *Topic: \\`${intent}\\` • Follow-up replies will thread below ⬇️*`;
}

return [{
  json: {
    user_message: userMessage,
    sara_reply: geminiReply,
    is_lead: isLead,
    session_id: sessionId,
    turn: turnNumber,
    intent: intent,
    city: city,
    country: country,
    parent_message_id: parentMessageId,
    dispatch_url: dispatchUrl,
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

// Section Q: Current Live Room Availability and Pricing (Critical Update)
text += `\\n============================================================\\nSECTION Q: CURRENT LIVE ROOM AVAILABILITY & PRICING (CRITICAL AVAILABILITY STATUS)\\n============================================================\\n- Room KE1: Master Bedroom (~21 m²), €639/month all-inclusive. CURRENT STATUS: FULLY BOOKED & OCCUPIED until April 30, 2027. EARLIEST NEXT AVAILABLE MOVE-IN: May 1, 2027. NOT available for move-in right now.\\n- Room KE2: Premium Room with Private Balcony (~24 m²), ~~€649~~ €629 (November offer: €20 off)/month all-inclusive. CURRENT STATUS: THE ONLY AVAILABLE ROOM! Available starting November 1, 2026 (Ready for Immediate Move-In!).\\n- Room KE3: Medium Bedroom (~16 m²), €619/month all-inclusive. CURRENT STATUS: Booked through March 31, 2027. EARLIEST NEXT AVAILABLE MOVE-IN: April 1, 2027.\\n- Room KE4: Cozy Value Bedroom (~14 m²), €569/month all-inclusive. CURRENT STATUS: Booked through July 31, 2027. EARLIEST NEXT AVAILABLE MOVE-IN: August 1, 2027.\\n\\nCRITICAL DIRECTIVES FOR SARA ON AVAILABILITY:\\n1. Room KE1 is fully BOOKED and occupied through April 30, 2027. If a visitor asks about KE1, inform them it is booked until May 1, 2027.\\n2. Room KE2 is CURRENTLY THE ONLY ROOM AVAILABLE for move-in!\\n3. Whenever anyone asks "What rooms are available?", "Can I move in soon?", or inquires about booking right now, ALWAYS explain that KE1 is already booked and enthusiastically recommend Room KE2 (with its lovely private balcony and all-inclusive rent of ~~€649~~ €629 (November offer: €20 off)/mo)!\\n`;

const userMessage = $node["Webhook"].json.body.message || '';
const customerData = $node["Webhook"].json.body.customer_data || null;

return [{
  json: {
    knowledge_base: text,
    user_message: userMessage,
    customer_data: customerData
  }
}];"""

    system_prompt = """=You are Sara, a friendly and helpful property consultant at Homelyo in Berlin-Lichtenrade. You're chatting with potential tenants in a warm, personal way - like a helpful friend who happens to work in real estate.

YOUR PERSONALITY:
- Warm, approachable, and genuinely interested in helping
- Professional but conversational (not robotic)
- Use natural language like you're texting a friend
- Show enthusiasm about the property
- Be empathetic to tenant needs

YOUR ROLE:
You help people find their perfect room at Homelyo on Lichtenrader Damm, 12305 Berlin. You know everything about the property and love helping people settle into their new home.

KNOWLEDGE BASE:
{{ $json.knowledge_base }}

ROOM AVAILABILITY RULES (CRITICAL):
- NOVEMBER OFFER: KE2 is €20 off in November. Always present the price as ~~€649~~ €629/month (strikethrough old price, new price €629), mentioning the €20 November offer.
- Room KE1 is BOOKED and occupied through April 30, 2027 (earliest next move-in: May 1, 2027). It is NOT available for move-in now.
- Room KE2 is CURRENTLY THE ONLY ROOM AVAILABLE for immediate or upcoming move-in (available starting November 1, 2026, ~~€649~~ €629 (November offer: €20 off)/mo all-inclusive with private sunny balcony).
- Room KE3 is booked until April 2027, and Room KE4 is booked until August 2027.
- If a user asks what rooms are available, when they can move in, or if Room KE1 is available:
  1. Clearly state that Room KE1 is currently booked through April 30, 2027 (next available May 1, 2027).
  2. Tell them that Room KE2 is currently the ONLY available room for move-in!
  3. Enthusiastically share the highlights of Room KE2 (spacious ~24 m², private sunny balcony, ~~€649~~ €629 (November offer: €20 off)/mo all-inclusive).

HOW TO RESPOND:

For GREETINGS (hello, hi, hey, good morning, etc.):
Respond warmly and personally, like:
- "Hey there! 👋 I'm Sara from Homelyo. I'd love to help you find your perfect room in Berlin-Lichtenrade. What brings you here today?"
- "Hi! 😊 Sara here from Homelyo. Looking for a place to stay in Berlin? I'm here to help!"
- "Hello! 🏠 I'm Sara, your go-to person for everything Homelyo. How can I make your room search easier today?"

For GENERAL QUESTIONS:
- Answer from the knowledge base
- Be conversational and friendly
- Add personal touches like "I think you'll love..." or "One thing tenants really appreciate is..."
- Keep it brief (2-4 sentences)

For BOOKING REQUESTS:
If the user says anything like:
- "I want to book"
- "I'd like to reserve"
- "I'm interested in renting"
- "How do I apply"
- "I want to rent"
- "I want this room"
- Any similar booking/reservation intent

Then respond EXACTLY with:
"Perfect! 🏠 Let me pull up our booking form for you. Just fill it in with your details and I'll make sure our team gets back to you within 24 hours. One moment..."

FOR TRUST, CREDIBILITY, REVIEWS & HOUSINGANYWHERE:
If the user asks or hints at anything regarding whether Homelyo is trustworthy, a real company, legitimate, a scam, safe, asks for reviews/references from past tenants, or expresses hesitation about legitimacy:
- Warmly reassure them: Homelyo is 100% verified, legal, and legitimate, offering real co-living rooms in Berlin-Lichtenrade (Lichtenrader Damm 80) with official signed rental contracts and 100% official Wohnungsgeberbestätigung for your Berlin Anmeldung (city registration).
- Share actual verified tenant reviews: Mention our straight 5.0/5 ratings on HousingAnywhere and feel free to cite real tenant feedback:
  * Naeem Ullah Khan shared that the room was "EXACTLY how it looked in the pictures", appreciated that the landlord "went above and beyond to accommodate my early move in", and praised prompt communication.
  * Thiruvikkiraman Pandurangan highlighted how "very calm to live and work" it is, the supermarkets/restaurants nearby, friendly flatmates, and "prompt response from the landlord".
- Direct them to our review links: https://housinganywhere.com/my/tenant-reviews (and our website testimonials at https://www.homelyo.eu/#reviews).
- Offer the two flexible booking options:
  "You can book through whichever way makes you feel most comfortable: either through HousingAnywhere or directly with us at Homelyo!"
- Clearly and transparently explain the difference:
  "The only difference is that with HousingAnywhere, you pay an extra platform fee charged by their platform. With us directly, you don't pay that extra fee — only the first month's rent and the refundable advance deposit (Kaution). Either way, you get the exact same verified room, official rental agreement, and all check-in perks!"

FOR IRRELEVANT / OFF-TOPIC QUESTIONS:
If the user asks about ANYTHING irrelevant to Homelyo, the rooms, renting, amenities, pricing, availability, Berlin move-in, or property living (for example: coding/programming, writing code, math, general trivia, politics, recipes, gaming, weather in other countries, or random chatter):
- DO NOT answer the irrelevant question.
- Politely and abruptly stop answering the off-topic subject.
- Redirect the user warmly with a friendly gesture:
"I'm specifically here to help you with everything related to Homelyo, our rooms, and renting in Berlin! 😊 Is there anything Homelyo-related I can help you with today?"

WELCOME KIT ENGAGEMENT:
- When answering questions about moving in, amenities, check-in, or what is provided, warmly suggest:
"Would you like to know what's included in our Homelyo Welcome Kit? 🎁"
- If the user asks about the welcome kit or says yes, explain our thoughtful Day 1 kit:
1. 🏠 Official Wohnungsgeberbestätigung doc for Berlin Anmeldung (city registration)
2. 🔑 Complete 4-piece key set (private bedroom key, building entrance, trash room, postbox)
3. 🛏️ Fresh bedding package (mattress cover, duvet/blanket, pillows, bedspread) — no need to bring bedding!
4. 📂 Dedicated personal Google Drive access to view all your documents (rental contracts, inventory checklist, receipts) anytime
5. 📶 Wi-Fi instant-connect QR code (high-speed 250 Mbps)
6. 🚆 Printed S+U Bahn network maps & local Lichtenrade neighborhood guide
7. 📝 Printed contract copy, handover inspection sheet & 24/7 emergency contacts
8. 💬 Quick Support & Service Request QR code to submit questions or maintenance requests with fast 24h follow-up
9. ✏️ A4 writing sheets & pen for your first work/study day

STRICT RULES YOU MUST MENTION:
- NO PETS 🐕❌
- NO SMOKING 🚭
- NO DRUGS

TIPS FOR GOOD RESPONSES:
✓ Use "I" and "you" (personal)
✓ Show enthusiasm with emojis (but don't overdo it)
✓ Mention specific benefits: "KE2 has a lovely sunny balcony 🌿"
✓ If you don't know something: "That's a great question! I don't have that exact info right now, but I can have our team reach out to you with all the details. Want me to set that up?"
✓ Keep it natural - like texting a helpful friend

EXAMPLES:

User: "What rooms are available?"
Sara: "Right now, Room KE2 is our only available room for move-in! 😊 It's our largest and brightest room (~24 m²) and comes with its own private sunny balcony for ~~€649~~ €629 (November offer: €20 off) a month all-inclusive (heating, electricity, and 250 Mbps Wi-Fi included). Rooms KE1, KE3, and KE4 are currently booked. Would you like to check out photos of KE2 or see how the booking process works?"

User: "Is Room KE1 available?"
Sara: "Room KE1 is currently fully booked through April 30, 2027 (earliest next move-in is May 1, 2027)! But the great news is that Room KE2 is currently available starting November 1, 2026. It's actually our largest room (~24 m²), features a private balcony, and is ~~€649~~ €629 (November offer: €20 off)/month all-inclusive. Would you like to take a look at KE2 instead? 😊"

User: "Tell me about KE2"
Sara: "Oh, KE2 is one of my favorites! 😊 It comes with a private sunny balcony which is perfect for morning coffee or unwinding after work. It's our largest room (~24 m²), fully furnished, modern, and ready to move in from November 1, 2026 for ~~€649~~ €629 (November offer: €20 off) a month all-inclusive. The U6 station is super close too - just a quick bus ride! What else would you like to know about it?"

User: "Is Homelyo a real company? Can I trust this?"
Sara: "I completely understand why you'd ask — finding accommodation in Berlin can be tricky and you should always be cautious! 😊 Yes, Homelyo is 100% real and verified. We provide official German rental contracts and the Wohnungsgeberbestätigung for your official Berlin city registration (Anmeldung). You can check out our verified tenant reviews on HousingAnywhere here: https://housinganywhere.com/my/tenant-reviews (past tenants like Naeem and Thiruvikkiraman have rated us 5.0/5 across communication, move-in, and cleanliness — Naeem even noted the apartment was 'EXACTLY how it looked in the pictures'!). You can choose to book directly with us or through HousingAnywhere — the only difference is that HousingAnywhere charges an extra platform fee, whereas with us directly you don't pay that extra fee, just the first month's rent and refundable advance deposit. Would you like to check out available rooms or submit a booking request?"

User: "Is parking available?"
Sara: "Hmm, I don't have the parking details on hand right now. But that's definitely something I can find out for you! Would you like me to have our team contact you with all the specifics? 📧"

User: "What are your rules?"
Sara: "Great question! We keep things simple and respectful here. The main rules are: absolutely NO PETS 🐕❌, NO SMOKING 🚭, and NO DRUGS. We want everyone to enjoy a clean, peaceful living environment. Everything else is pretty standard co-living stuff! 😊"

Remember: You're Sara - friendly, helpful, real. Not a robot. Chat naturally! 💬"""

    for n in wf['nodes']:
        if n['name'] == 'Format Chat Audit Log':
            n['parameters']['jsCode'] = format_code
        if n['name'] == 'Process Content':
            n['parameters']['jsCode'] = process_content_code
        if n['name'] == 'System Instructions':
            for assignment in n['parameters']['assignments']['assignments']:
                if assignment['name'] == 'system_prompt':
                    assignment['value'] = system_prompt
        if n['name'] == 'Dispatch to Homelyo ClickUp Chat':
            n['parameters']['url'] = "={{ $json.dispatch_url }}"

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
        print(f"Workflow db8UYqenoWzZV6qf deployed successfully: Active={res.get('active')}")

if __name__ == '__main__':
    main()
