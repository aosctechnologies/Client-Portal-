import os
import json
import requests
import faiss
import numpy as np
import re
import json
# from fastapi import FastAPI, UploadFile, HTTPException
from dotenv import load_dotenv
from PyPDF2 import PdfReader
import docx
from fastapi import FastAPI, UploadFile, HTTPException, File


# ================= ENV =================

load_dotenv(dotenv_path=".env", override=True)

OR_API_KEY = os.getenv("OPENROUTER_API_KEY")

if not OR_API_KEY:
    raise RuntimeError("OPENROUTER_API_KEY not found in environment")

# ================= FASTAPI =================

app = FastAPI(title="AI Document Validation Agent (OpenRouter)")

# ================= TEXT EXTRACTION =================

def extract_text(file: UploadFile) -> str:
    if file.filename.lower().endswith(".pdf"):
        reader = PdfReader(file.file)
        return "\n".join([page.extract_text() or "" for page in reader.pages])

    if file.filename.lower().endswith(".docx"):
        doc = docx.Document(file.file)
        return "\n".join([p.text for p in doc.paragraphs])

    raise HTTPException(status_code=400, detail="Unsupported file format")

#======================Helpers=======================

def parse_ai_json(ai_text: str):
    """
    Extract and parse JSON from AI response safely
    """
    if not ai_text:
        return {"error": "Empty AI response"}

    # Remove ```json ``` or ``` wrappers
    cleaned = re.sub(r"```json|```", "", ai_text).strip()

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        return {
            "error": "Invalid JSON returned by AI",
            "raw_response": ai_text
        }

# ================= CHUNKING =================

def chunk_text(text, chunk_size=800, overlap=100):
    chunks = []
    start = 0
    while start < len(text):
        end = start + chunk_size
        chunks.append(text[start:end])
        start = end - overlap
    return chunks

# ================= EMBEDDINGS (OPENROUTER) =================

def embed_chunks(chunks):
    embeddings = []

    headers = {
        "Authorization": f"Bearer {OR_API_KEY}",
        "Content-Type": "application/json"
    }

    for chunk in chunks:
        payload = {
            "model": "text-embedding-3-small",
            "input": chunk
        }

        res = requests.post(
            "https://openrouter.ai/api/v1/embeddings",
            headers=headers,
            json=payload,
            timeout=30
        )

        if res.status_code != 200:
            raise RuntimeError("Embedding generation failed")

        embeddings.append({
            "text": chunk,
            "embedding": res.json()["data"][0]["embedding"]
        })

    return embeddings

# ================= VECTOR STORE =================

def build_vector_store(embedded_chunks):
    dim = len(embedded_chunks[0]["embedding"])
    index = faiss.IndexFlatL2(dim)

    vectors = np.array(
        [e["embedding"] for e in embedded_chunks],
        dtype="float32"
    )

    index.add(vectors)
    texts = [e["text"] for e in embedded_chunks]

    return index, texts

def semantic_search(index, texts, query, top_k=3):
    headers = {
        "Authorization": f"Bearer {OR_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": "text-embedding-3-small",
        "input": query
    }

    res = requests.post(
        "https://openrouter.ai/api/v1/embeddings",
        headers=headers,
        json=payload,
        timeout=30
    )

    if res.status_code != 200:
        raise RuntimeError("Query embedding failed")

    query_embedding = res.json()["data"][0]["embedding"]
    query_vector = np.array([query_embedding], dtype="float32")

    _, indices = index.search(query_vector, top_k)
    return [texts[i] for i in indices[0]]

# ================= AI ANALYSIS (PROMPT KEPT HERE) =================

def analyze_document_with_ai(context: str):
    headers = {
        "Authorization": f"Bearer {OR_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": "http://localhost:8000",
        "X-Title": "doc-validation-agent"
    }

    prompt = f"""
You are an intelligent business document analysis and validation assistant.

Your task is to analyze the provided document context and do the following:

1. Identify what type of document this appears to be 
   (e.g., business compliance document, registration certificate, invoice, contract, policy, letter, or other).

2. Based on the identified document type, determine the key expected fields or information
   that such a document should normally contain
   (for example: business name, registration numbers, dates, identifiers, signatures, addresses, etc.).

3. Check which of those expected fields are PRESENT in the document
   and which are MISSING or UNCLEAR.

4. If the document is incomplete or missing important information,
   clearly list the missing fields.

5. If the document is not a standard business/compliance document,
   still analyze it and explain what information is present and what may be missing or unclear.

6. Highlight any potential risks, inconsistencies, or concerns based on the document content.

Document Context:
{context}

Return STRICT JSON:
{{
 "document_type": "string",
  "detected_fields": {{
    "field_name": "extracted value or null"
}}
 "missing_fields": ["field_name"],
  "risks": ["string"],
  "summary": "brief, clear explanation of the document status"
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
        return {"error": "AI analysis unavailable",
                "status_code": res.status_code,
                "response": res.text}

    try:
        return res.json()["choices"][0]["message"]["content"]
    except Exception:
        return {"error": "AI analysis failed"}

# ================= API ENDPOINT =================

@app.post("/validate-document")
async def validate_document(file: UploadFile = File(...)):
    text = extract_text(file)

    if not text.strip():
        raise HTTPException(status_code=400, detail="No text extracted from document")

    chunks = chunk_text(text)
    embedded_chunks = embed_chunks(chunks)

    index, texts = build_vector_store(embedded_chunks)

    relevant_chunks = semantic_search(
        index,
        texts,
        query="Australian Business Number ABN document date compliance",
        top_k=3
    )
    MAX_CONTEXT_CHARS = 2500
    context = "\n".join(relevant_chunks)
    context = context[:MAX_CONTEXT_CHARS]
    
    ai_raw = analyze_document_with_ai(context)
    ai_parsed = parse_ai_json(ai_raw)

    missing = ai_parsed.get("missing_fields", [])
    risks = ai_parsed.get("risks", [])
    invalid = []
    status = "CLEAR"
    if missing or invalid or risks:
        status = "NEEDS_ATTENTION"

    return {
    "document_type": ai_parsed.get("document_type", "Unknown"),
    "status": status,
    "issues": {
        "missing_fields": missing,
        "invalid_fields": invalid,
        "risks": risks
    },
    "message": ai_parsed.get(
        "summary",
        "Document analysis completed."
    )
}
    
    # ai_parsed = parse_ai_json(ai_raw)

    # return {
    #     "filename": file.filename,
    #     "ai_validation": ai_parsed
    # }
