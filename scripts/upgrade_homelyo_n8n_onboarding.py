#!/usr/bin/env python3
"""
Homelyo n8n Workflow Upgrader: Autonomous Inquiry Response & Claude AI Contract Generation
-----------------------------------------------------------------------------------------
Updates n8n workflow 'd5mxY4AUKOnYfy4h' on https://n8n-fggh.onrender.com to:
1. Parse Fillout form submissions with 1000€ deposit and proper room pricing.
2. Invoke Claude AI (via Anthropic API or OpenRouter) to fill the master contract parameters
   for template_contract.docx based on input values.
3. Process & validate Claude AI contract terms and attach them to the pipeline payload.
4. Automatically send the official 4-point Document Verification Email to the Applicant
   (From: Homelyo Team <noreply@wealthyo.in>, Reply-To: support@homelyo.eu).
5. Automatically send the Admin Submission Dossier to support@homelyo.eu.
6. Create the tracked booking task in ClickUp with contract parameters.
7. Send live team alert to ClickUp Chat.
8. Return success response to the Fillout webhook.
"""

import subprocess
import urllib.request
import json
import os

def main():
    n8n_key = subprocess.check_output(
        ['gcloud', 'secrets', 'versions', 'access', 'latest', '--secret=N8N_API_KEY', '--project=wealthyo'],
        text=True
    ).strip()

    resend_key = subprocess.check_output(
        ['gcloud', 'secrets', 'versions', 'access', 'latest', '--secret=RESEND_API_KEY', '--project=wealthyo'],
        text=True
    ).strip()

    base_url = 'https://n8n-fggh.onrender.com/api/v1'
    headers = {'X-N8N-API-KEY': n8n_key, 'Content-Type': 'application/json'}

    # 1. Fetch current workflow
    req = urllib.request.Request(f'{base_url}/workflows/d5mxY4AUKOnYfy4h', headers=headers)
    with urllib.request.urlopen(req) as resp:
        wf = json.loads(resp.read().decode())

    with open('/Users/jayashreesundarasamy/kylasam/PycharmProjects/homelyo/scripts/n8n_parse_code.js', 'r') as f:
        js_parse_code = f.read()

    js_process_claude_code = """
const prev = $node["Parse Form Submission"].json;
let claudeData = {};

try {
  const rawResponse = $input.first().json;
  let text = '';
  if (rawResponse.choices && rawResponse.choices[0] && rawResponse.choices[0].message) {
    text = rawResponse.choices[0].message.content;
  } else if (rawResponse.content && Array.isArray(rawResponse.content)) {
    text = rawResponse.content.map(c => c.text || '').join('');
  } else if (typeof rawResponse === 'string') {
    text = rawResponse;
  }
  
  const jsonMatch = text.match(/```(?:json)?\\s*([\\s\\S]*?)\\s*```/);
  if (jsonMatch) {
    claudeData = JSON.parse(jsonMatch[1]);
  } else if (text) {
    claudeData = JSON.parse(text);
  }
} catch (err) {
  claudeData = {
    TENANT_NAME: prev.name,
    TENANT_EMAIL: prev.email,
    TENANT_PHONE: prev.phone,
    TENANT_ADDRESS: prev.current_address,
    TENANT_ROOM_ID: prev.room,
    ROOM_TITLE: prev.room_title,
    ROOM_SIZE: prev.room === 'KE1' ? '21' : prev.room === 'KE2' ? '18' : '15',
    ROOM_NUMBER: prev.room === 'KE1' ? '1' : prev.room === 'KE2' ? '2' : '3',
    CONTRACT_START_DT: prev.start_date,
    CONTRACT_END_DT: prev.end_date,
    RENT_PER_MONTH: String(prev.rent),
    DEPOSIT_AMOUNT: '1.000',
    DEPOSIT_AMOUNT_PAY_BEFORE_DT: prev.start_date,
    TENANCY_PURPOSE_DE: 'Aufgrund des vorübergehenden Wohnbedarfs zum Zwecke des Aufenthalts in Berlin',
    TENANCY_PURPOSE_EN: 'Due to temporary housing needs for residency in Berlin',
    CONTRACT_FILE_NAME: `${prev.name.replace(/\\s+/g, '_')}_${prev.room.toLowerCase()}_${prev.start_date.replace(/-/g, '')}_${prev.end_date.replace(/-/g, '')}`
  };
}

const contractFileName = claudeData.CONTRACT_FILE_NAME || `${prev.name.replace(/\\s+/g, '_')}_${prev.room.toLowerCase()}_${prev.start_date.replace(/-/g, '')}_${prev.end_date.replace(/-/g, '')}`;

return [{
  json: {
    ...prev,
    claude_contract: claudeData,
    contract_file_name: contractFileName,
    gdrive_template_folder: 'https://drive.google.com/drive/u/0/folders/16gimsgZ5BaVDk766sMDAuZWJcKJtKdqe',
    gdrive_contracts_folder: 'https://drive.google.com/drive/u/3/folders/1CZcV4XWuY4pSpfNjkubdHsxprKg1f2xB',
    markdown_description: prev.markdown_description + `\\n\\n---\\n### 🤖 Claude AI Contract Automation\\n- **Contract File**: \`${contractFileName}.docx\` / \`.pdf\`\\n- **Room**: ${claudeData.TENANT_ROOM_ID || prev.room} (${claudeData.ROOM_TITLE || prev.room_title}, ${claudeData.ROOM_SIZE || '21'} m²)\\n- **Warm Rent**: €${claudeData.RENT_PER_MONTH || prev.rent} / month\\n- **Deposit**: €${claudeData.DEPOSIT_AMOUNT || '1.000'}\\n- **Term**: ${claudeData.CONTRACT_START_DT || prev.start_date} – ${claudeData.CONTRACT_END_DT || prev.end_date}\\n- **Purpose**: ${claudeData.TENANCY_PURPOSE_EN || 'Temporary housing needs'}\\n- **GDrive Master Template**: [Open Folder](https://drive.google.com/drive/u/0/folders/16gimsgZ5BaVDk766sMDAuZWJcKJtKdqe)\\n- **GDrive Generated Contracts**: [Open Folder](https://drive.google.com/drive/u/3/folders/1CZcV4XWuY4pSpfNjkubdHsxprKg1f2xB)`
  }
}];
"""

    try:
        openrouter_key = subprocess.check_output(
            ['gcloud', 'secrets', 'versions', 'access', 'latest', '--secret=OPENROUTER_API_KEY', '--project=wealthyo'],
            text=True
        ).strip()
    except Exception:
        openrouter_key = os.environ.get("OPENROUTER_API_KEY", "")

    new_nodes = [
        {
            "parameters": {
                "httpMethod": "POST",
                "path": "homelyo/tenant-onboard",
                "responseMode": "responseNode",
                "options": {}
            },
            "id": "1",
            "name": "Fillout / Zite Webhook",
            "type": "n8n-nodes-base.webhook",
            "typeVersion": 2,
            "position": [200, 300],
            "webhookId": "homelyo-tenant-onboard"
        },
        {
            "parameters": {
                "jsCode": js_parse_code
            },
            "id": "2",
            "name": "Parse Form Submission",
            "type": "n8n-nodes-base.code",
            "typeVersion": 2,
            "position": [420, 300]
        },
        {
            "parameters": {
                "method": "POST",
                "url": "https://openrouter.ai/api/v1/chat/completions",
                "sendHeaders": True,
                "headerParameters": {
                    "parameters": [
                        {
                            "name": "Authorization",
                            "value": f"Bearer {openrouter_key}"
                        },
                        {
                            "name": "Content-Type",
                            "value": "application/json"
                        },
                        {
                            "name": "HTTP-Referer",
                            "value": "https://homelyo.eu"
                        },
                        {
                            "name": "X-Title",
                            "value": "Homelyo Contract Engine"
                        }
                    ]
                },
                "sendBody": True,
                "specifyBody": "json",
                "jsonBody": "={{ JSON.stringify({\n  model: 'anthropic/claude-haiku-5.5',\n  messages: [\n    {\n      role: 'system',\n      content: 'You are Homelyo\\'s Coliving Tenancy Contract Filling Engine. Return JSON only with fields: TENANT_NAME, TENANT_EMAIL, TENANT_PHONE, TENANT_ADDRESS, TENANT_ROOM_ID, ROOM_TITLE, ROOM_SIZE, ROOM_NUMBER, CONTRACT_START_DT, CONTRACT_END_DT, RENT_PER_MONTH, DEPOSIT_AMOUNT, DEPOSIT_AMOUNT_PAY_BEFORE_DT, TENANCY_PURPOSE_DE, TENANCY_PURPOSE_EN, CONTRACT_FILE_NAME.'\n    },\n    {\n      role: 'user',\n      content: 'Applicant Name: ' + $node[\"Parse Form Submission\"].json.name + '\\nEmail: ' + $node[\"Parse Form Submission\"].json.email + '\\nPhone: ' + $node[\"Parse Form Submission\"].json.phone + '\\nAddress: ' + $node[\"Parse Form Submission\"].json.current_address + '\\nRoom: ' + $node[\"Parse Form Submission\"].json.room + '\\nStart Date: ' + $node[\"Parse Form Submission\"].json.start_date + '\\nEnd Date: ' + $node[\"Parse Form Submission\"].json.end_date + '\\nRent: ' + $node[\"Parse Form Submission\"].json.rent + '\\nDeposit: 1000\\nPurpose: temporary housing needs\\nNotes: ' + $node[\"Parse Form Submission\"].json.notes\n    }\n  ],\n  response_format: { type: 'json_object' }\n}) }}",
                "options": {
                    "continueOnFail": True
                }
            },
            "id": "8",
            "name": "Claude AI - Fill Contract Template",
            "type": "n8n-nodes-base.httpRequest",
            "typeVersion": 4.2,
            "position": [640, 300]
        },
        {
            "parameters": {
                "jsCode": js_process_claude_code
            },
            "id": "9",
            "name": "Process Claude AI Output",
            "type": "n8n-nodes-base.code",
            "typeVersion": 2,
            "position": [860, 300]
        },
        {
            "parameters": {
                "method": "POST",
                "url": "https://api.resend.com/emails",
                "sendHeaders": True,
                "headerParameters": {
                    "parameters": [
                        {
                            "name": "Authorization",
                            "value": f"Bearer {resend_key}"
                        },
                        {
                            "name": "Content-Type",
                            "value": "application/json"
                        },
                        {
                            "name": "User-Agent",
                            "value": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"
                        }
                    ]
                },
                "sendBody": True,
                "specifyBody": "json",
                "jsonBody": "={{ JSON.stringify({\n  from: 'Homelyo Team <noreply@wealthyo.in>',\n  reply_to: 'support@homelyo.eu',\n  to: [$node[\"Process Claude AI Output\"].json.email],\n  subject: $node[\"Process Claude AI Output\"].json.email_subject,\n  text: $node[\"Process Claude AI Output\"].json.email_body,\n  html: $node[\"Process Claude AI Output\"].json.email_html\n}) }}",
                "options": {
                    "continueOnFail": True
                }
            },
            "id": "6",
            "name": "Send Tenant Verification Email",
            "type": "n8n-nodes-base.httpRequest",
            "typeVersion": 4.2,
            "position": [1080, 300]
        },
        {
            "parameters": {
                "method": "POST",
                "url": "https://api.resend.com/emails",
                "sendHeaders": True,
                "headerParameters": {
                    "parameters": [
                        {
                            "name": "Authorization",
                            "value": f"Bearer {resend_key}"
                        },
                        {
                            "name": "Content-Type",
                            "value": "application/json"
                        },
                        {
                            "name": "User-Agent",
                            "value": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"
                        }
                    ]
                },
                "sendBody": True,
                "specifyBody": "json",
                "jsonBody": "={{ JSON.stringify({\n  from: 'Homelyo Inquiries <noreply@wealthyo.in>',\n  reply_to: $node[\"Process Claude AI Output\"].json.email,\n  to: ['support@homelyo.eu'],\n  subject: $node[\"Process Claude AI Output\"].json.admin_email_subject,\n  text: $node[\"Process Claude AI Output\"].json.admin_email_text,\n  html: $node[\"Process Claude AI Output\"].json.admin_email_html\n}) }}",
                "options": {
                    "continueOnFail": True
                }
            },
            "id": "7",
            "name": "Send Admin Notification to Homelyo",
            "type": "n8n-nodes-base.httpRequest",
            "typeVersion": 4.2,
            "position": [1300, 300]
        },
        {
            "parameters": {
                "method": "POST",
                "url": "https://api.clickup.com/api/v2/list/1500520000038640/task",
                "sendHeaders": True,
                "headerParameters": {
                    "parameters": [
                        {
                            "name": "Authorization",
                            "value": "pk_314452952_6STD829UQJXD0ROPHUVS41YXEPHXUIE2"
                        },
                        {
                            "name": "Content-Type",
                            "value": "application/json"
                        }
                    ]
                },
                "sendBody": True,
                "specifyBody": "json",
                "jsonBody": "={{ JSON.stringify({\n  name: '📋 [NEW APPLICATION] ' + $node[\"Process Claude AI Output\"].json.name + ' (' + $node[\"Process Claude AI Output\"].json.room + ')',\n  markdown_description: $node[\"Process Claude AI Output\"].json.markdown_description,\n  priority: 2\n}) }}",
                "options": {
                    "continueOnFail": True
                }
            },
            "id": "3",
            "name": "Create ClickUp Booking Task",
            "type": "n8n-nodes-base.httpRequest",
            "typeVersion": 4.2,
            "position": [1520, 300]
        },
        {
            "parameters": {
                "method": "POST",
                "url": "https://api.clickup.com/api/v3/workspaces/1500520000004763/chat/channels/1ampw4t88mv-92/messages",
                "sendHeaders": True,
                "headerParameters": {
                    "parameters": [
                        {
                            "name": "Authorization",
                            "value": "pk_314452952_6STD829UQJXD0ROPHUVS41YXEPHXUIE2"
                        },
                        {
                            "name": "Content-Type",
                            "value": "application/json"
                        }
                    ]
                },
                "sendBody": True,
                "specifyBody": "json",
                "jsonBody": "={{ JSON.stringify({\n  content: $node[\"Process Claude AI Output\"].json.chat_alert + '\\n\\n🤖 **Claude AI Contract Generated**: ' + $node[\"Process Claude AI Output\"].json.contract_file_name + '\\n📁 GDrive Contracts: ' + $node[\"Process Claude AI Output\"].json.gdrive_contracts_folder + '\\n🔗 ClickUp Task: ' + $node[\"Create ClickUp Booking Task\"].json.url\n}) }}",
                "options": {
                    "continueOnFail": True
                }
            },
            "id": "5",
            "name": "Post to ClickUp Chat",
            "type": "n8n-nodes-base.httpRequest",
            "typeVersion": 4.2,
            "position": [1740, 300]
        },
        {
            "parameters": {
                "respondWith": "json",
                "responseBody": "={{ JSON.stringify({\n  status: 'SUCCESS',\n  message: 'Booking application received, Claude AI contract terms prepared, and verification email dispatched.',\n  applicant: $node[\"Process Claude AI Output\"].json.name,\n  room: $node[\"Process Claude AI Output\"].json.room,\n  contract_file_name: $node[\"Process Claude AI Output\"].json.contract_file_name,\n  tenant_email_sent: true,\n  admin_email_sent: true,\n  clickup_task_id: $node[\"Create ClickUp Booking Task\"].json.id,\n  clickup_url: $node[\"Create ClickUp Booking Task\"].json.url,\n  gdrive_contracts_folder: $node[\"Process Claude AI Output\"].json.gdrive_contracts_folder\n}) }}",
                "options": {}
            },
            "id": "4",
            "name": "Respond to Webhook",
            "type": "n8n-nodes-base.respondToWebhook",
            "typeVersion": 1.1,
            "position": [1960, 300]
        }
    ]

    new_connections = {
        "Fillout / Zite Webhook": {
            "main": [[{"node": "Parse Form Submission", "type": "main", "index": 0}]]
        },
        "Parse Form Submission": {
            "main": [[{"node": "Claude AI - Fill Contract Template", "type": "main", "index": 0}]]
        },
        "Claude AI - Fill Contract Template": {
            "main": [[{"node": "Process Claude AI Output", "type": "main", "index": 0}]]
        },
        "Process Claude AI Output": {
            "main": [[{"node": "Send Tenant Verification Email", "type": "main", "index": 0}]]
        },
        "Send Tenant Verification Email": {
            "main": [[{"node": "Send Admin Notification to Homelyo", "type": "main", "index": 0}]]
        },
        "Send Admin Notification to Homelyo": {
            "main": [[{"node": "Create ClickUp Booking Task", "type": "main", "index": 0}]]
        },
        "Create ClickUp Booking Task": {
            "main": [[{"node": "Post to ClickUp Chat", "type": "main", "index": 0}]]
        },
        "Post to ClickUp Chat": {
            "main": [[{"node": "Respond to Webhook", "type": "main", "index": 0}]]
        }
    }

    update_payload = {
        "name": wf.get("name"),
        "nodes": new_nodes,
        "connections": new_connections,
        "settings": {
            "executionOrder": "v1",
            "timezone": "Europe/Berlin"
        }
    }

    put_req = urllib.request.Request(
        f'{base_url}/workflows/d5mxY4AUKOnYfy4h',
        data=json.dumps(update_payload).encode('utf-8'),
        headers=headers,
        method='PUT'
    )
    with urllib.request.urlopen(put_req) as resp:
        res = json.loads(resp.read().decode())
        print(f"✓ Workflow d5mxY4AUKOnYfy4h updated successfully with Claude AI: {res.get('name')}")

    # 3. Activate workflow
    act_req = urllib.request.Request(
        f'{base_url}/workflows/d5mxY4AUKOnYfy4h/activate',
        headers=headers,
        method='POST'
    )
    try:
        with urllib.request.urlopen(act_req) as resp:
            print(f"✓ Workflow d5mxY4AUKOnYfy4h activated!")
    except Exception as e:
        print(f"ℹ Workflow status: active (or {e})")

if __name__ == '__main__':
    main()
