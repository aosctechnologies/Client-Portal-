import os
import json
import requests
import traceback
from fastapi import FastAPI, HTTPException, Request, Query
from fastapi.responses import JSONResponse
from dotenv import load_dotenv

# ================= ENV =================
load_dotenv(dotenv_path=".env", override=True)

OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
if not OPENROUTER_API_KEY:
    raise RuntimeError("OPENROUTER_API_KEY not found")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

app = FastAPI(title="AI Onboarding Agent - KYC Validator")

# ================= AUTH =================
def get_headers(request: Request):
    auth = request.headers.get("Authorization")
    if not auth or not auth.startswith("Bearer "):
        return None
    return {
        "Authorization": auth,
        "Content-Type": "application/json",
        "Accept": "application/json"
    }

# ================= HELPER: FIELDS → TEXT =================
def fields_to_text(fields: dict) -> str:
    lines = []
    for key, value in fields.items():
        clean_key = key.replace("_x0020_", " ").replace("_x002f_", "/")
        if value and str(value).strip():
            lines.append(f"{clean_key}: {value}")
        else:
            lines.append(f"{clean_key}: Missing")
    return "\n".join(lines)

# ================= AI ANALYSIS =================
def analyze_onboarding_with_ai(context: str):
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": "http://localhost:8000",
        "X-Title": "ai-onboarding-agent"
    }

    prompt = f"""
You are an intelligent client onboarding and KYC validation assistant.

Analyze the onboarding information below and do the following:

1. Identify whether this data represents a client onboarding / KYC record.
2. Determine which important onboarding or KYC fields are missing.
3. Identify any invalid, unclear, or suspicious information.
4. Highlight any compliance or verification risks.
5. Decide if the onboarding can be considered COMPLETE or NEEDS_ATTENTION.

ONLY report issues. Do NOT repeat fields that are already valid.

Onboarding Data:
{context}

Return STRICT JSON in this format:

{{
  "document_type": "Client Onboarding / KYC",
  "status": "CLEAR or NEEDS_ATTENTION",
  "issues": {{
    "missing_fields": ["field name"],
    "invalid_fields": ["field name or issue"],
    "risks": ["risk description"]
  }},
  "message": "short plain-English explanation"
}}
"""

    payload = {
        "model": "openai/gpt-4o-mini",
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2,
        "max_tokens": 300
    }

    res = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers=headers,
        json=payload,
        timeout=60
    )

    if res.status_code != 200:
        return {
            "error": "AI analysis failed",
            "status_code": res.status_code,
            "response": res.text
        }

    return res.json()["choices"][0]["message"]["content"]

# ================= AI JSON PARSER =================
def parse_ai_json(ai_text: str):
    try:
        cleaned = ai_text.replace("```json", "").replace("```", "").strip()
        return json.loads(cleaned)
    except Exception:
        return {
            "document_type": "Client Onboarding / KYC",
            "status": "NEEDS_ATTENTION",
            "issues": {
                "missing_fields": [],
                "invalid_fields": [],
                "risks": ["AI response could not be parsed"]
            },
            "message": "Unable to reliably analyze onboarding data."
        }

# ================= MAIN API =================
@app.get("/onboarding/{hostname}/sites/{site_name}/lists/{list_id}")
async def process_onboarding(
    request: Request,
    hostname: str,
    site_name: str,
    list_id: str,
    query: str = Query(..., description="Value to identify onboarding record")
):
    headers = get_headers(request)
    if not headers:
        raise HTTPException(status_code=401, detail="Bearer token missing")

    try:
        # 1. Resolve Site ID
        site_url = f"{GRAPH_BASE}/sites/{hostname}:/sites/{site_name}"
        site_res = requests.get(site_url, headers=headers)
        if not site_res.ok:
            return JSONResponse(status_code=site_res.status_code, content={"error": "Invalid site path"})

        site_id = site_res.json().get("id")

        # 2. Fetch list items
        list_url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?expand=fields"
        items_res = requests.get(list_url, headers=headers)
        items = items_res.json().get("value", [])

        target_item = None
        clean_query = query.strip().lower()

        # 3. Find matching record
        for item in items:
            fields = item.get("fields", {})
            for _, val in fields.items():
                if val and str(val).strip().lower() == clean_query:
                    target_item = item
                    break
            if target_item:
                break

        if not target_item:
            raise HTTPException(status_code=404, detail=f"No onboarding record found for '{query}'")

        fields = target_item.get("fields", {})

        # 4. Convert fields to AI-readable text
        context = fields_to_text(fields)

        # 5. AI validation
        # ai_raw = analyze_onboarding_with_ai(context)
        # ai_result = parse_ai_json(ai_raw)

        # return ai_result


        ai_raw = analyze_onboarding_with_ai(context)
        ai_result = parse_ai_json(ai_raw)
        issues = ai_result.get("issues", {})
        # Filter out SharePoint system fields
         
        SYSTEM_FIELDS = [
        "Attachments",
        "Edit",
        "_ComplianceFlags",
        "_ComplianceTag",
        "_ComplianceTagWrittenTime",
        "_ComplianceTagUserId"
        ]
        missing = [
            f for f in issues.get("missing_fields", [])
            if f not in SYSTEM_FIELDS
]

        invalid = issues.get("invalid_fields", [])
        risks = issues.get("risks", [])

        messages = []
        if missing:
            # messages.append(
            messages.append(
                f"Missing required fields: {', '.join(missing)}"
            )
        if invalid:
            messages.append(
                f"Invalid fields detected: {', '.join(invalid)}"
            )
            
        if risks:
            messages.append(
        f"Risk indicators found: {', '.join(risks)}")
            
        final_message = " | ".join(messages) if messages else "The onboarding record is complete and valid."


        return {
              "missing_fields": missing,
              "invalid_fields": invalid,
              "risk_fields": risks,
              "message": final_message

            #   "message": ai_result.get(
            #       "message",
            #       "The onboarding record requires attention."
    
}


    except Exception as e:
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)






