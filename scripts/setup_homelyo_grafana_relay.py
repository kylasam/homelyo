#!/usr/bin/env python3
"""
Deploy Homelyo Grafana Alert Relay to n8n & configure Grafana Alerting.
Routes all Homelyo alerts to ClickUp Chat channel 1ampw4t88mv-112.
"""

import json
import subprocess
import urllib.request
import urllib.error

def get_secret(name):
    return subprocess.check_output(
        ['gcloud', 'secrets', 'versions', 'access', 'latest', f'--secret={name}', '--project=wealthyo']
    ).decode().strip()

def deploy_n8n_workflow():
    n8n_key = get_secret('N8N_API_KEY')
    clickup_token = get_secret('HOMELYO_CLICKUP_API_TOKEN')

    workflow_json = {
        'name': 'Homelyo - Grafana Alerts to ClickUp Chat Relay',
        'settings': {
            'executionOrder': 'v1',
            'timezone': 'Europe/Berlin',
            'saveDataSuccessExecution': 'none',
            'saveDataErrorExecution': 'all'
        },
        'nodes': [
            {
                'parameters': {
                    'httpMethod': 'POST',
                    'path': 'homelyo/grafana-alerts',
                    'responseMode': 'onReceived',
                    'responseData': 'OK',
                    'options': {}
                },
                'name': 'Webhook: Grafana Alert Ingest',
                'type': 'n8n-nodes-base.webhook',
                'typeVersion': 2,
                'position': [100, 300],
                'webhookId': 'homelyo-grafana-alerts-relay'
            },
            {
                'parameters': {
                    'jsCode': '''const body = $input.first().json.body || $input.first().json;
const globalStatus = body.status || 'firing';
const alerts = body.alerts && body.alerts.length > 0 ? body.alerts : [{
    status: globalStatus,
    labels: body.commonLabels || {},
    annotations: body.commonAnnotations || {},
    startsAt: new Date().toISOString(),
    generatorURL: body.externalURL || 'https://kylash.grafana.net/alerting'
}];

const results = [];

for (const alert of alerts) {
    const isFiring = (alert.status || globalStatus) === 'firing';
    const icon = isFiring ? '🚨' : '✅';
    const statusText = isFiring ? 'FIRING / ACTIVE' : 'RESOLVED';
    const labels = alert.labels || {};
    const annotations = alert.annotations || {};

    const alertName = annotations.summary || labels.pipeline || labels.alertname || 'Homelyo Service Alert';
    const description = annotations.description || 'No detailed description provided.';
    const startsAt = alert.startsAt ? new Date(alert.startsAt).toLocaleString('en-GB', { timeZone: 'Europe/Berlin' }) : 'N/A';
    const endsAt = (alert.endsAt && alert.endsAt !== '0001-01-01T00:00:00Z') ? new Date(alert.endsAt).toLocaleString('en-GB', { timeZone: 'Europe/Berlin' }) : null;
    const generatorURL = alert.generatorURL || body.externalURL || 'https://kylash.grafana.net/alerting';

    let markdown = `${icon} **[${statusText}] ${alertName}**\\n\\n`;
    markdown += `• **Status:** \\`${statusText}\\`\\n`;
    if (labels.pipeline) markdown += `• **Pipeline:** ${labels.pipeline}\\n`;
    if (labels.workflow_id) markdown += `• **Workflow ID:** \\`${labels.workflow_id}\\`\\n`;
    if (labels.latest_execution_id) {
        markdown += `• **Execution ID:** [${labels.latest_execution_id}](https://n8n-fggh.onrender.com/workflow/${labels.workflow_id}/executions/${labels.latest_execution_id})\\n`;
    }
    markdown += `• **Triggered At:** ${startsAt} CET\\n`;
    if (endsAt) markdown += `• **Resolved At:** ${endsAt} CET\\n`;
    markdown += `\\n**📋 Alert Details:**\\n${description}\\n\\n`;
    markdown += `🔗 [View in Grafana](${generatorURL})`;

    results.push({
        json: {
            content_format: 'text/md',
            content: markdown
        }
    });
}

return results;'''
                },
                'name': 'Format Alert Markdown',
                'type': 'n8n-nodes-base.code',
                'typeVersion': 2,
                'position': [350, 300]
            },
            {
                'parameters': {
                    'method': 'POST',
                    'url': 'https://api.clickup.com/api/v3/workspaces/1500520000004763/chat/channels/1ampw4t88mv-112/messages',
                    'sendHeaders': True,
                    'headerParameters': {
                        'parameters': [
                            {
                                'name': 'Authorization',
                                'value': clickup_token
                            },
                            {
                                'name': 'Content-Type',
                                'value': 'application/json'
                            }
                        ]
                    },
                    'sendBody': True,
                    'specifyBody': 'json',
                    'jsonBody': '={{ JSON.stringify({ content_format: $json.content_format, content: $json.content }) }}',
                    'options': {}
                },
                'name': 'ClickUp: Post to Grafana Alerts Chat',
                'type': 'n8n-nodes-base.httpRequest',
                'typeVersion': 4.2,
                'position': [600, 300]
            }
        ],
        'connections': {
            'Webhook: Grafana Alert Ingest': {
                'main': [
                    [
                        {
                            'node': 'Format Alert Markdown',
                            'type': 'main',
                            'index': 0
                        }
                    ]
                ]
            },
            'Format Alert Markdown': {
                'main': [
                    [
                        {
                            'node': 'ClickUp: Post to Grafana Alerts Chat',
                            'type': 'main',
                            'index': 0
                        }
                    ]
                ]
            }
        }
    }

    # Check if workflow already exists
    req_list = urllib.request.Request('https://n8n-fggh.onrender.com/api/v1/workflows', headers={'X-N8N-API-KEY': n8n_key})
    with urllib.request.urlopen(req_list) as resp:
        wfs = json.loads(resp.read().decode()).get('data', [])
    
    existing_wf = next((w for w in wfs if w['name'] == workflow_json['name']), None)
    if existing_wf:
        wf_id = existing_wf['id']
        update_url = f'https://n8n-fggh.onrender.com/api/v1/workflows/{wf_id}'
        req = urllib.request.Request(
            update_url,
            data=json.dumps(workflow_json).encode(),
            headers={'X-N8N-API-KEY': n8n_key, 'Content-Type': 'application/json'},
            method='PUT'
        )
        with urllib.request.urlopen(req) as resp:
            print(f"Updated existing workflow: {wf_id}")
    else:
        req = urllib.request.Request(
            'https://n8n-fggh.onrender.com/api/v1/workflows',
            data=json.dumps(workflow_json).encode(),
            headers={'X-N8N-API-KEY': n8n_key, 'Content-Type': 'application/json'},
            method='POST'
        )
        with urllib.request.urlopen(req) as resp:
            wf = json.loads(resp.read().decode())
            wf_id = wf['id']
            print(f"Created new workflow: {wf_id}")

    # Activate workflow
    act_req = urllib.request.Request(
        f'https://n8n-fggh.onrender.com/api/v1/workflows/{wf_id}/activate',
        headers={'X-N8N-API-KEY': n8n_key},
        method='POST'
    )
    with urllib.request.urlopen(act_req) as resp:
        print(f"Activated workflow {wf_id}: status {resp.status}")

    return wf_id

if __name__ == "__main__":
    deploy_n8n_workflow()
