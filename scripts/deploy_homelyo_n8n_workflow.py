import subprocess
import urllib.request
import json
import os

def main():
    n8n_key = subprocess.check_output(
        ['gcloud', 'secrets', 'versions', 'access', 'latest', '--secret=N8N_API_KEY', '--project=wealthyo'],
        text=True
    ).strip()

    homelyo_token = subprocess.check_output(
        ['gcloud', 'secrets', 'versions', 'access', 'latest', '--secret=HOMELYO_CLICKUP_API_TOKEN', '--project=wealthyo'],
        text=True
    ).strip()

    base_url = 'https://n8n-fggh.onrender.com/api/v1'
    headers = {'X-N8N-API-KEY': n8n_key, 'Content-Type': 'application/json'}

    # 1. Fetch current workflow
    req = urllib.request.Request(f'{base_url}/workflows/db8UYqenoWzZV6qf', headers=headers)
    with urllib.request.urlopen(req) as resp:
        wf = json.loads(resp.read().decode())

    # Format Chat Audit Log code
    format_code = """
const webhookData = $node["Webhook"].json.body || {};
const headers = $node["Webhook"].json.headers || {};
const geminiReply = $node["Gemini Flash API"].json.candidates?.[0]?.content?.parts?.[0]?.text || "No reply generated";
const userMessage = webhookData.message || "(empty message)";
const isLead = geminiReply.toLowerCase().includes("booking form") || geminiReply.toLowerCase().includes("booking request") || geminiReply.toLowerCase().includes("https://forms.fillout.com");

const country = headers["cf-ipcountry"] || headers["x-country"] || "DE";
const city = headers["cf-ipcity"] || "Berlin";
const userAgent = headers["user-agent"] || "Web Browser";

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

const leadBadge = isLead ? "🔥 **[QUALIFIED BOOKING LEAD]**" : "💬 **[LIVE CHAT INQUIRY]**";

const chatMarkdown = `${leadBadge}
• **Location:** ${city}, ${country}
• **Topic / Intent:** \`${intent}\`
• **Device / Browser:** ${userAgent.substring(0, 40)}

👤 **Visitor Message:**
> ${userMessage.replace(/\\n/g, '\\n> ')}

🤖 **Sara AI Reply:**
> ${geminiReply.substring(0, 1000).replace(/\\n/g, '\\n> ')}
`;

return [{
  json: {
    user_message: userMessage,
    sara_reply: geminiReply,
    is_lead: isLead,
    intent: intent,
    city: city,
    country: country,
    homelyo_chat_content: chatMarkdown
  }
}];
"""

    format_node = {
        'id': 'format-chat-audit-log-id',
        'name': 'Format Chat Audit Log',
        'type': 'n8n-nodes-base.code',
        'typeVersion': 2,
        'position': [2048, 160],
        'parameters': {
            'jsCode': format_code
        }
    }

    clickup_chat_node = {
        'id': 'dispatch-homelyo-clickup-chat-id',
        'name': 'Dispatch to Homelyo ClickUp Chat',
        'type': 'n8n-nodes-base.httpRequest',
        'typeVersion': 4.2,
        'position': [2272, 160],
        'parameters': {
            'method': 'POST',
            'url': 'https://api.clickup.com/api/v3/workspaces/1500520000004763/chat/channels/1ampw4t88mv-92/messages',
            'sendHeaders': True,
            'specifyHeaders': 'keypair',
            'headerParameters': {
                'parameters': [
                    {'name': 'Authorization', 'value': homelyo_token},
                    {'name': 'Content-Type', 'value': 'application/json'},
                    {'name': 'accept', 'value': 'application/json'}
                ]
            },
            'sendBody': True,
            'contentType': 'json',
            'specifyBody': 'json',
            'jsonBody': '={{ JSON.stringify({ type: "message", content_format: "text/md", content: $json.homelyo_chat_content }) }}',
            'options': {}
        }
    }

    # Only retain core Sara processing nodes
    keep_node_names = ['Webhook', 'Get Knowledge Base', 'Process Content', 'System Instructions', 'Gemini Flash API', 'Respond (Normal)']
    clean_nodes = [n for n in wf.get('nodes', []) if n['name'] in keep_node_names]
    clean_nodes.append(format_node)
    clean_nodes.append(clickup_chat_node)

    new_connections = {
        'Webhook': {
            'main': [[{'node': 'Get Knowledge Base', 'type': 'main', 'index': 0}]]
        },
        'Get Knowledge Base': {
            'main': [[{'node': 'Process Content', 'type': 'main', 'index': 0}]]
        },
        'Process Content': {
            'main': [[{'node': 'System Instructions', 'type': 'main', 'index': 0}]]
        },
        'System Instructions': {
            'main': [[{'node': 'Gemini Flash API', 'type': 'main', 'index': 0}]]
        },
        'Gemini Flash API': {
            'main': [
                [
                    {'node': 'Respond (Normal)', 'type': 'main', 'index': 0},
                    {'node': 'Format Chat Audit Log', 'type': 'main', 'index': 0}
                ]
            ]
        },
        'Format Chat Audit Log': {
            'main': [
                [{'node': 'Dispatch to Homelyo ClickUp Chat', 'type': 'main', 'index': 0}]
            ]
        }
    }

    update_payload = {
        'name': 'Homelyo Chatbot with Forms',
        'nodes': clean_nodes,
        'connections': new_connections,
        'settings': wf.get('settings', {})
    }

    put_req = urllib.request.Request(
        f'{base_url}/workflows/db8UYqenoWzZV6qf',
        data=json.dumps(update_payload).encode('utf-8'),
        headers=headers,
        method='PUT'
    )
    with urllib.request.urlopen(put_req) as put_resp:
        res = json.loads(put_resp.read().decode())
        print('✅ Workflow updated successfully! ID:', res.get('id'), 'Active:', res.get('active'))

    # Activate workflow
    act_req = urllib.request.Request(f'{base_url}/workflows/db8UYqenoWzZV6qf/activate', headers=headers, method='POST')
    try:
        with urllib.request.urlopen(act_req) as act_resp:
            res = json.loads(act_resp.read().decode())
            print('✅ Workflow activated successfully! Active:', res.get('active'))
    except urllib.error.HTTPError as e:
        print('Activation HTTP status:', e.code)

if __name__ == '__main__':
    main()
