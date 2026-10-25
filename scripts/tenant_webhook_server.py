#!/usr/bin/env python3
"""
Homelyo Tenant Webhook Server & Zite / Fillout Adapter
------------------------------------------------------
Listens for incoming booking applications / inquiries from:
1. Zite / Fillout Webhook (https://app.zite.com/editor/gq1XFEqEWJus/integrations)
2. Website book.html form
3. Direct JSON API call

Extracts fields, triggers auto_onboard_tenant pipeline, and responds with
Google Drive folder URL, contract Google Doc URL, and PDF download link.
"""

import sys
import os
import json
import re
from http.server import HTTPServer, BaseHTTPRequestHandler
from auto_onboard_tenant import onboard_tenant

PORT = int(os.environ.get("PORT", 8787))

def extract_fields(body):
    """
    Robust extractor handling both flat JSON and Fillout/Zite questions array.
    """
    fields = {}
    
    # 1. Check if payload has Fillout 'questions' array
    if isinstance(body, dict) and "questions" in body and isinstance(body["questions"], list):
        for q in body["questions"]:
            name = str(q.get("name", "")).strip().lower()
            val = q.get("value")
            fields[name] = val
    elif isinstance(body, dict):
        for k, v in body.items():
            fields[str(k).strip().lower()] = v

    # Helper pick
    def pick(*keys, default=""):
        for k in keys:
            k_low = k.lower()
            for fk, fv in fields.items():
                if k_low in fk and fv:
                    return str(fv).strip()
        return default

    name = pick("full name", "applicant name", "name", default="Guest Applicant")
    email = pick("email address", "email", default="support@homelyo.eu")
    phone = pick("phone number", "whatsapp", "phone", default="")
    raw_room = pick("selected room", "room preference", "room", default="KE1")
    
    # Clean room to KE1, KE2, KE3, KE4
    m = re.search(r"KE[1-4]", raw_room, re.IGNORECASE)
    room = m.group(0).upper() if m else "KE1"
    
    start_date = pick("target move-in", "move-in date", "move in date", "start date", "start", default="2026-11-01")
    # Clean date format if DD/MM/YYYY or similar
    if "/" in start_date:
        parts = start_date.split("/")
        if len(parts) == 3:
            if len(parts[2]) == 4:
                start_date = f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
    
    end_date = pick("end date", "end", default="")
    raw_duration = pick("stay length", "duration", "months", default="6")
    try:
        duration_months = int(re.search(r"\d+", raw_duration).group(0))
    except Exception:
        duration_months = 6

    platform = pick("platform", "source", default="Zite / Fillout Form gq1XFEqEWJus")

    return {
        "name": name,
        "email": email,
        "phone": phone,
        "room": room,
        "start_date": start_date,
        "end_date": end_date or None,
        "duration_months": duration_months,
        "platform": platform
    }

class WebhookHandler(BaseHTTPRequestHandler):
    def _send_response(self, status, payload):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.end_headers()
        self.wfile.write(json.dumps(payload, indent=2).encode("utf-8"))

    def do_OPTIONS(self):
        self._send_response(200, {"status": "OK"})

    def do_GET(self):
        if self.path == "/health" or self.path == "/":
            self._send_response(200, {
                "status": "UP",
                "service": "Homelyo Autonomous Onboarding & Contract Engine",
                "version": "1.0.0",
                "fillout_form_id": "gq1XFEqEWJus",
                "master_gdrive_folder": "1jVYbBfrm3SHqIou-Xf9iVrYaLsro5X0p"
            })
        else:
            self._send_response(404, {"error": "Not Found"})

    def do_POST(self):
        content_length = int(self.headers.get("Content-Length", 0))
        raw_body = self.rfile.read(content_length).decode("utf-8")
        
        try:
            body = json.loads(raw_body)
        except Exception as e:
            self._send_response(400, {"error": f"Invalid JSON payload: {e}"})
            return

        try:
            params = extract_fields(body)
            print(f"📥 Received Webhook Submission: {params['name']} ({params['room']}, {params['start_date']})")
            
            # Execute Onboarding Pipeline
            result = onboard_tenant(
                name=params["name"],
                email=params["email"],
                phone=params["phone"],
                room=params["room"],
                start_date=params["start_date"],
                end_date=params["end_date"],
                duration_months=params["duration_months"],
                platform=params["platform"]
            )
            self._send_response(200, result)
        except Exception as e:
            print(f"❌ Error during onboarding execution: {e}")
            self._send_response(500, {"status": "ERROR", "message": str(e)})

def run_server():
    server_address = ("", PORT)
    httpd = HTTPServer(server_address, WebhookHandler)
    print(f"⚡ Homelyo Tenant Onboarding Webhook Server running on port {PORT}...")
    print(f"👉 Webhook endpoint: http://localhost:{PORT}/api/onboard-tenant")
    print(f"👉 Health check: http://localhost:{PORT}/health")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server.")
        httpd.server_close()

if __name__ == "__main__":
    run_server()
