import subprocess
import urllib.request
import json
import os

def main():
    n8n_key = subprocess.check_output(
        ['gcloud', 'secrets', 'versions', 'access', 'latest', '--secret=N8N_API_KEY', '--project=wealthyo'],
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
            "position": [240, 300],
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
            "position": [460, 300]
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
                "jsonBody": "={{ JSON.stringify({\n  name: '📋 [NEW APPLICATION] ' + $json.name + ' (' + $json.room + ')',\n  markdown_description: $json.markdown_description,\n  priority: 2\n}) }}",
                "options": {}
            },
            "id": "3",
            "name": "Create ClickUp Booking Task",
            "type": "n8n-nodes-base.httpRequest",
            "typeVersion": 4.2,
            "position": [700, 300]
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
                "jsonBody": "={{ JSON.stringify({\n  content: $node[\"Parse Form Submission\"].json.chat_alert + '\\n\\n🔗 ClickUp Task: ' + $node[\"Create ClickUp Booking Task\"].json.url\n}) }}",
                "options": {
                    "continueOnFail": True
                }
            },
            "id": "5",
            "name": "Post to ClickUp Chat",
            "type": "n8n-nodes-base.httpRequest",
            "typeVersion": 4.2,
            "position": [920, 300]
        },
        {
            "parameters": {
                "respondWith": "json",
                "responseBody": "={{ JSON.stringify({\n  status: 'SUCCESS',\n  message: 'Booking application received and processed.',\n  applicant: $node[\"Parse Form Submission\"].json.name,\n  room: $node[\"Parse Form Submission\"].json.room,\n  clickup_task_id: $node[\"Create ClickUp Booking Task\"].json.id,\n  clickup_url: $node[\"Create ClickUp Booking Task\"].json.url\n}) }}",
                "options": {}
            },
            "id": "4",
            "name": "Respond to Webhook",
            "type": "n8n-nodes-base.respondToWebhook",
            "typeVersion": 1.1,
            "position": [1140, 300]
        }
    ]

    new_connections = {
        "Fillout / Zite Webhook": {
            "main": [[{"node": "Parse Form Submission", "type": "main", "index": 0}]]
        },
        "Parse Form Submission": {
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
        print(f"Workflow d5mxY4AUKOnYfy4h updated successfully: {res.get('name')}")

if __name__ == '__main__':
    main()
