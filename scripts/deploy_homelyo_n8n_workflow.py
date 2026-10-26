import subprocess
import urllib.request
import json

def main():
    n8n_key = subprocess.check_output(
        ['gcloud', 'secrets', 'versions', 'access', 'latest', '--secret=N8N_API_KEY', '--project=wealthyo'],
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

const leadHeader = isLead ? "🔥 **[HIGH-INTENT BOOKING LEAD]**\\n" : "";

const chatMarkdown = `${leadHeader}💬 **Visitor [Thread #${sessionId} • ${turn}]**
> *"${userMessage.replace(/\\n/g, ' ')}"*

🤖 **Sara AI Concierge:**
${geminiReply}

📍 *Visitor: ${city}, ${country} • Topic: \`${intent}\`*`;

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

    for node in wf["nodes"]:
        if node["name"] == "Format Chat Audit Log":
            node["parameters"]["jsCode"] = format_code
        if node["name"] == "Dispatch to Homelyo ClickUp Chat":
            node["parameters"]["url"] = "https://api.clickup.com/api/v3/workspaces/1500520000004763/chat/channels/1ampw4t88mv-92/messages"
            node["parameters"]["headerParameters"]["parameters"] = [
                {"name": "Authorization", "value": "pk_105682594_CCTOZ8K23KHPSAY3JQTVV1G23QTN63IG"},
                {"name": "Content-Type", "value": "application/json"},
                {"name": "accept", "value": "application/json"}
            ]

    update_payload = {
        "name": wf["name"],
        "nodes": wf["nodes"],
        "connections": wf["connections"],
        "settings": wf.get("settings", {})
    }

    put_req = urllib.request.Request(
        f'{base_url}/workflows/db8UYqenoWzZV6qf',
        data=json.dumps(update_payload).encode('utf-8'),
        headers=headers,
        method='PUT'
    )
    with urllib.request.urlopen(put_req) as resp:
        res = json.loads(resp.read().decode())
        print(f"Workflow db8UYqenoWzZV6qf updated successfully: {res.get('name')}")

if __name__ == '__main__':
    main()
