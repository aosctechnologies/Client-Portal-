# import requests
# from flask import Flask, jsonify, request

# app = Flask(__name__)

# GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# # SharePoint site details
# HOSTNAME = "aoscaustralia.sharepoint.com"
# SITE_PATH = "/sites/CPA"

# # --------------------------------------------------
# # Helper: Extract Bearer token
# # --------------------------------------------------
# def get_headers():
#     auth = request.headers.get("Authorization")
#     if not auth or not auth.startswith("Bearer "):
#         raise Exception("Authorization header missing or invalid")

#     return {
#         "Authorization": auth,
#         "Content-Type": "application/json"
#     }

# # --------------------------------------------------
# # STEP 1: Resolve Site ID
# # --------------------------------------------------
# @app.route("/sharepoint/site", methods=["GET"])
# def get_site():
#     try:
#         headers = get_headers()

#         url = f"{GRAPH_BASE}/sites/{HOSTNAME}:{SITE_PATH}"
#         res = requests.get(url, headers=headers, timeout=30)
#         res.raise_for_status()

#         site = res.json()

#         return jsonify({
#             "site_id": site["id"],
#             "displayName": site["displayName"],
#             "webUrl": site["webUrl"]
#         })

#     except Exception as e:
#         return jsonify({"error": str(e)}), 403

# # --------------------------------------------------
# # STEP 2A: Get SharePoint Lists
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists", methods=["GET"])
# def get_lists(site_id):
#     try:
#         headers = get_headers()

#         url = f"{GRAPH_BASE}/sites/{site_id}/lists"
#         res = requests.get(url, headers=headers, timeout=30)
#         res.raise_for_status()

#         lists = []
#         for lst in res.json().get("value", []):
#             lists.append({
#                 "list_id": lst["id"],
#                 "name": lst["name"],
#                 "displayName": lst.get("displayName"),
#                 "type": lst.get("list", {}).get("template")
#             })

#         return jsonify(lists)

#     except Exception as e:
#         return jsonify({"error": str(e)}), 403

# # --------------------------------------------------
# # STEP 3A: Get List Items
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["GET"])
# def get_list_items(site_id, list_id):
#     try:
#         headers = get_headers()

#         url = (
#             f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
#             "?expand=fields"
#         )
#         res = requests.get(url, headers=headers, timeout=30)
#         res.raise_for_status()

#         return jsonify(res.json().get("value", []))

#     except Exception as e:
#         return jsonify({"error": str(e)}), 403

# # --------------------------------------------------
# # STEP 2B: Get Document Libraries
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries", methods=["GET"])
# def get_libraries(site_id):
#     try:
#         headers = get_headers()

#         url = f"{GRAPH_BASE}/sites/{site_id}/drives"
#         res = requests.get(url, headers=headers, timeout=30)
#         res.raise_for_status()

#         libraries = []
#         for drive in res.json().get("value", []):
#             if drive.get("driveType") == "documentLibrary":
#                 libraries.append({
#                     "drive_id": drive["id"],
#                     "name": drive["name"],
#                     "webUrl": drive["webUrl"]
#                 })

#         return jsonify(libraries)

#     except Exception as e:
#         return jsonify({"error": str(e)}), 403

# # --------------------------------------------------
# # STEP 3B: Get Documents from Library
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries/<drive_id>/documents", methods=["GET"])
# def get_documents(site_id, drive_id):
#     try:
#         headers = get_headers()

#         url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"
#         res = requests.get(url, headers=headers, timeout=30)
#         res.raise_for_status()

#         return jsonify(res.json().get("value", []))

#     except Exception as e:
#         return jsonify({"error": str(e)}), 403

# # --------------------------------------------------
# # STEP 4B: Upload Document to Library
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries/<drive_id>/upload", methods=["POST"])
# def upload_document(site_id, drive_id):
#     try:
#         if "file" not in request.files:
#             return jsonify({"error": "No file provided"}), 400

#         file = request.files["file"]
#         headers = get_headers()

#         upload_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{file.filename}:/content"

#         upload_headers = {
#             "Authorization": headers["Authorization"],
#             "Content-Type": file.content_type or "application/octet-stream"
#         }

#         res = requests.put(
#             upload_url,
#             headers=upload_headers,
#             data=file.read(),
#             timeout=60
#         )
#         res.raise_for_status()

#         uploaded = res.json()

#         return jsonify({
#             "uploaded_file": {
#                 "id": uploaded["id"],
#                 "name": uploaded["name"],
#                 "webUrl": uploaded["webUrl"]
#             }
#         })

#     except Exception as e:
#         return jsonify({"error": str(e)}), 403

# # --------------------------------------------------
# # Health
# # --------------------------------------------------
# @app.route("/health", methods=["GET"])
# def health():
#     return jsonify({"status": "SharePoint Lists & Libraries API running"})

# # --------------------------------------------------
# # Run
# # --------------------------------------------------
# if __name__ == "__main__":
#     print("🚀 Backend running on http://localhost:5050")
#     app.run(host="127.0.0.1", port=5050, debug=False)
































# import requests
# from flask import Flask, jsonify, request

# app = Flask(__name__)

# GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# HOSTNAME = "aoscaustralia.sharepoint.com"
# SITE_PATH = "/sites/CPA"

# # --------------------------------------------------
# # Helper: Get Delegated User Token
# # --------------------------------------------------
# def get_headers():
#     auth = request.headers.get("Authorization")
#     if not auth or not auth.startswith("Bearer "):
#         return None

#     return {
#         "Authorization": auth,
#         "Content-Type": "application/json"
#     }

# # --------------------------------------------------
# # STEP 1: Resolve Site
# # --------------------------------------------------
# @app.route("/sharepoint/site", methods=["GET"])
# def get_site():
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{HOSTNAME}:{SITE_PATH}"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     site = res.json()
#     return jsonify({
#         "site_id": site["id"],
#         "displayName": site["displayName"],
#         "webUrl": site["webUrl"]
#     })

# # --------------------------------------------------
# # STEP 2: Get Lists
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists", methods=["GET"])
# def get_lists(site_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     lists = [{
#         "list_id": lst["id"],
#         "displayName": lst.get("displayName"),
#         "template": lst.get("list", {}).get("template")
#     } for lst in res.json().get("value", [])]

#     return jsonify(lists)

# # --------------------------------------------------
# # STEP 3: Get List Items
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["GET"])
# def get_list_items(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?expand=fields"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json().get("value", []))

# # --------------------------------------------------
# # STEP 4: Create List Item (POST) – DELEGATED
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["POST"])
# def create_list_item(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     payload = request.get_json()
#     if not payload or "fields" not in payload:
#         return jsonify({"error": "fields object is required"}), 400

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"

#     res = requests.post(
#         url,
#         headers=headers,
#         json={"fields": payload["fields"]}
#     )

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json()), 201

# # --------------------------------------------------
# # STEP 5: Get Document Libraries
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries", methods=["GET"])
# def get_libraries(site_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/drives"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     libraries = [
#         d for d in res.json().get("value", [])
#         if d.get("driveType") == "documentLibrary"
#     ]

#     return jsonify(libraries)

# # --------------------------------------------------
# # STEP 6: Get Documents
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries/<drive_id>/documents", methods=["GET"])
# def get_documents(site_id, drive_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json().get("value", []))

# # --------------------------------------------------
# # STEP 7: Upload Document
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries/<drive_id>/upload", methods=["POST"])
# def upload_document(site_id, drive_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     if "file" not in request.files:
#         return jsonify({"error": "No file provided"}), 400

#     file = request.files["file"]

#     upload_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{file.filename}:/content"

#     upload_headers = {
#         "Authorization": headers["Authorization"],
#         "Content-Type": file.content_type or "application/octet-stream"
#     }

#     res = requests.put(upload_url, headers=upload_headers, data=file.read())

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json())

# # --------------------------------------------------
# # STEP 8: Get Users (People Picker)
# # --------------------------------------------------
# @app.route("/graph/users", methods=["GET"])
# def get_users():
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/users?$select=displayName,mail,userPrincipalName"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     users = []
#     for u in res.json().get("value", []):
#         email = u.get("mail") or u.get("userPrincipalName")
#         users.append({
#             "displayName": u["displayName"],
#             "email": email,
#             "claims": f"i:0#.f|membership|{email}"
#         })

#     return jsonify(users)

# # --------------------------------------------------
# # HEALTH
# # --------------------------------------------------
# @app.route("/health", methods=["GET"])
# def health():
#     return jsonify({"status": "Delegated SharePoint API running"})

# # --------------------------------------------------
# # RUN
# # --------------------------------------------------
# if __name__ == "__main__":
#     print("🚀 Backend running on http://localhost:5050")
#     app.run(host="127.0.0.1", port=5050, debug=False)






















































# import requests
# from flask import Flask, jsonify, request

# app = Flask(__name__)

# GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# HOSTNAME = "aoscaustralia.sharepoint.com"
# SITE_PATH = "/sites/CPA"

# # --------------------------------------------------
# # Helper: Get Delegated User Token
# # --------------------------------------------------
# def get_headers():
#     auth = request.headers.get("Authorization")
#     if not auth or not auth.startswith("Bearer "):
#         return None

#     return {
#         "Authorization": auth,
#         "Content-Type": "application/json"
#     }

# # --------------------------------------------------
# # STEP 1: Resolve SharePoint Site
# # --------------------------------------------------
# @app.route("/sharepoint/site", methods=["GET"])
# def get_site():
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{HOSTNAME}:{SITE_PATH}"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     site = res.json()
#     return jsonify({
#         "site_id": site["id"],
#         "displayName": site["displayName"],
#         "webUrl": site["webUrl"]
#     })

# # --------------------------------------------------
# # STEP 2: Get Lists
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists", methods=["GET"])
# def get_lists(site_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     lists = [{
#         "list_id": lst["id"],
#         "displayName": lst.get("displayName"),
#         "template": lst.get("list", {}).get("template")
#     } for lst in res.json().get("value", [])]

#     return jsonify(lists)

# # --------------------------------------------------
# # STEP 3: Get List Items (READ)
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["GET"])
# def get_list_items(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?expand=fields"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json().get("value", []))

# # --------------------------------------------------
# # STEP 4: Create List Item (POST)
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["POST"])
# def create_list_item(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     payload = request.get_json()
#     if not payload or "fields" not in payload:
#         return jsonify({"error": "fields object required"}), 400

#     fields = payload["fields"]

#     # 🚫 Block calculated/system fields
#     forbidden = ["PercentComplete", "LastUpdated"]
#     for f in forbidden:
#         fields.pop(f, None)

#     # ✅ Validate AssignedStaff
#     if "AssignedStaff" in fields:
#         staff = fields["AssignedStaff"]
#         if not isinstance(staff, dict) or "claims" not in staff:
#             return jsonify({"error": "AssignedStaff must be a claims object"}), 400

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"

#     res = requests.post(
#         url,
#         headers=headers,
#         json={"fields": fields}
#     )

#     if not res.ok:
#         return jsonify({
#             "error": "Graph API error",
#             "details": res.text
#         }), res.status_code

#     return jsonify(res.json()), 201


# # --------------------------------------------------
# # STEP 5: Get Document Libraries
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries", methods=["GET"])
# def get_libraries(site_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/drives"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     libraries = [
#         d for d in res.json().get("value", [])
#         if d.get("driveType") == "documentLibrary"
#     ]

#     return jsonify(libraries)

# # --------------------------------------------------
# # STEP 6: Get Documents from Library
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries/<drive_id>/documents", methods=["GET"])
# def get_documents(site_id, drive_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json().get("value", []))

# # --------------------------------------------------
# # STEP 7: Upload Document
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries/<drive_id>/upload", methods=["POST"])
# def upload_document(site_id, drive_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     if "file" not in request.files:
#         return jsonify({"error": "No file provided"}), 400

#     file = request.files["file"]
#     upload_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{file.filename}:/content"

#     upload_headers = {
#         "Authorization": headers["Authorization"],
#         "Content-Type": file.content_type or "application/octet-stream"
#     }

#     res = requests.put(upload_url, headers=upload_headers, data=file.read())

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json())

# # --------------------------------------------------
# # STEP 8: Get Users (People Picker)
# # --------------------------------------------------
# @app.route("/graph/users", methods=["GET"])
# def graph_get_users():
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/users?$select=displayName,mail,userPrincipalName"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     users = []
#     for u in res.json().get("value", []):
#         email = u.get("mail") or u.get("userPrincipalName")
#         if not email:
#             continue

#         users.append({
#             "displayName": u["displayName"],
#             "email": email,
#             "claims": f"i:0#.f|membership|{email}"
#         })

#     return jsonify(users)

# # --------------------------------------------------
# # HEALTH
# # --------------------------------------------------
# @app.route("/health", methods=["GET"])
# def health():
#     return jsonify({"status": "Delegated SharePoint API running"})

# # --------------------------------------------------
# # RUN
# # --------------------------------------------------
# if __name__ == "__main__":
#     print("🚀 Backend running on http://localhost:5050")
#     app.run(host="127.0.0.1", port=5050, debug=False)




















































# import requests
# from flask import Flask, jsonify, request

# app = Flask(__name__)

# GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# HOSTNAME = "aoscaustralia.sharepoint.com"
# SITE_PATH = "/sites/CPA"

# # --------------------------------------------------
# # Helper: Get Delegated User Token
# # --------------------------------------------------
# def get_headers():
#     auth = request.headers.get("Authorization")
#     if not auth or not auth.startswith("Bearer "):
#         return None

#     return {
#         "Authorization": auth,
#         "Content-Type": "application/json"
#     }

# # --------------------------------------------------
# # STEP 1: Resolve SharePoint Site
# # --------------------------------------------------
# @app.route("/sharepoint/site", methods=["GET"])
# def get_site():
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{HOSTNAME}:{SITE_PATH}"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     site = res.json()
#     return jsonify({
#         "site_id": site["id"],
#         "displayName": site["displayName"],
#         "webUrl": site["webUrl"]
#     })

# # --------------------------------------------------
# # STEP 2: Get Lists
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists", methods=["GET"])
# def get_lists(site_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     lists = [{
#         "list_id": lst["id"],
#         "displayName": lst.get("displayName"),
#         "template": lst.get("list", {}).get("template")
#     } for lst in res.json().get("value", [])]

#     return jsonify(lists)

# # --------------------------------------------------
# # STEP 3: Get List Items
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["GET"])
# def get_list_items(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?expand=fields"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json().get("value", []))

# # --------------------------------------------------
# # STEP 4: Create List Item (POST) – FULLY FIXED
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["POST"])
# def create_list_item(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     body = request.get_json()
#     if not body or "fields" not in body:
#         return jsonify({"error": "fields object required"}), 400

#     fields = body["fields"]

#     # 🚫 Remove system / calculated fields
#     for f in ["PercentComplete", "LastUpdated", "Created", "Modified"]:
#         fields.pop(f, None)

#     # ✅ AssignedStaff – SINGLE select only
#     if "AssignedStaff" in fields:
#         staff = fields["AssignedStaff"]

#         if not isinstance(staff, dict) or "claims" not in staff:
#             return jsonify({
#                 "error": "AssignedStaff must be a single claims object"
#             }), 400

#     print("POSTING FIELDS TO GRAPH:", fields)

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
#     res = requests.post(url, headers=headers, json={"fields": fields})

#     if not res.ok:
#         return jsonify({
#             "error": "Graph API error",
#             "details": res.text
#         }), res.status_code

#     return jsonify(res.json()), 201

# @app.route("/sharepoint/<site_id>/lists/<list_id>/columns", methods=["GET"])
# def get_list_columns(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/columns"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({
#             "error": "Graph API error",
#             "details": res.text
#         }), res.status_code

#     return jsonify(res.json())


# # --------------------------------------------------
# # STEP 5: Get Document Libraries
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries", methods=["GET"])
# def get_libraries(site_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/drives"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     libraries = [
#         d for d in res.json().get("value", [])
#         if d.get("driveType") == "documentLibrary"
#     ]

#     return jsonify(libraries)

# # --------------------------------------------------
# # STEP 6: Get Documents
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries/<drive_id>/documents", methods=["GET"])
# def get_documents(site_id, drive_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json().get("value", []))

# # --------------------------------------------------
# # STEP 7: Upload Document
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries/<drive_id>/upload", methods=["POST"])
# def upload_document(site_id, drive_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     if "file" not in request.files:
#         return jsonify({"error": "No file provided"}), 400

#     file = request.files["file"]

#     upload_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{file.filename}:/content"

#     upload_headers = {
#         "Authorization": headers["Authorization"],
#         "Content-Type": file.content_type or "application/octet-stream"
#     }

#     res = requests.put(upload_url, headers=upload_headers, data=file.read())

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json())

# # --------------------------------------------------
# # STEP 8: Get Users (People Picker)
# # --------------------------------------------------
# @app.route("/graph/users", methods=["GET"])
# def graph_get_users():
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/users?$select=displayName,mail,userPrincipalName"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     users = []
#     for u in res.json().get("value", []):
#         email = u.get("mail") or u.get("userPrincipalName")
#         if not email:
#             continue

#         users.append({
#             "displayName": u["displayName"],
#             "email": email,
#             "claims": f"i:0#.f|membership|{email}"
#         })

#     return jsonify(users)

# # --------------------------------------------------
# # HEALTH
# # --------------------------------------------------
# @app.route("/health", methods=["GET"])
# def health():
#     return jsonify({"status": "Delegated SharePoint API running"})

# # --------------------------------------------------
# # RUN
# # --------------------------------------------------
# if __name__ == "__main__":
#     print("🚀 Backend running on http://localhost:5050")
#     app.run(host="127.0.0.1", port=5050, debug=False)





# import requests
# from flask import Flask, jsonify, request

# app = Flask(__name__)

# GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# HOSTNAME = "aoscaustralia.sharepoint.com"
# SITE_PATH = "/sites/CPA"

# # --------------------------------------------------
# # Helper: Extract Delegated Bearer Token
# # --------------------------------------------------
# def get_headers():
#     auth = request.headers.get("Authorization")
#     if not auth or not auth.startswith("Bearer "):
#         return None

#     return {
#         "Authorization": auth,
#         "Content-Type": "application/json"
#     }

# # --------------------------------------------------
# # STEP 1: Resolve SharePoint Site
# # --------------------------------------------------
# @app.route("/sharepoint/site", methods=["GET"])
# def get_site():
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{HOSTNAME}:{SITE_PATH}"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     site = res.json()
#     return jsonify({
#         "site_id": site["id"],
#         "displayName": site["displayName"],
#         "webUrl": site["webUrl"]
        
#     })

# # --------------------------------------------------
# # STEP 2: Get SharePoint Lists
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists", methods=["GET"])
# def get_lists(site_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify([
#         {
#             "list_id": lst["id"],
#             "displayName": lst.get("displayName"),
#             "template": lst.get("list", {}).get("template")
#         }
#         for lst in res.json().get("value", [])
#     ])

# # --------------------------------------------------
# # STEP 3: Get List Items
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["GET"])
# def get_list_items(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?expand=fields"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return jsonify(res.json().get("value", []))

# # --------------------------------------------------
# # 🔥 STEP 4: UPSERT PROFILE (CREATE OR UPDATE BY EMAIL)
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists/<list_id>/profile", methods=["POST"])
# def upsert_profile(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     body = request.get_json()
#     if not body or "email" not in body:
#         return jsonify({"error": "email is required"}), 400

#     email = body["email"]

#     # 1️⃣ Search by Email
#     search_url = (
#         f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
#         f"?expand=fields&$filter=fields/Email eq '{email}'"
#     )

#     search_res = requests.get(search_url, headers=headers)
#     if not search_res.ok:
#         return jsonify({"error": "Search failed", "details": search_res.text}), 500

#     items = search_res.json().get("value", [])

#     # Map frontend → SharePoint columns
#     fields = {
#         "DisplayName": body.get("displayName"),
#         "LegalName": body.get("legalName"),
#         "Email": email,
#         "Phone": body.get("phone"),
#         "DOB": body.get("dob"),
#         "AddressLine1": body.get("addressLine1"),
#         "AddressLine2": body.get("addressLine2"),
#         "City": body.get("city"),
#         "State": body.get("state"),
#         "Zip": body.get("zip"),
#         "Country": body.get("country"),
#         "ClientId": body.get("clientId"),
#     }

#     # Remove empty values
#     fields = {k: v for k, v in fields.items() if v is not None}

#     # 2️⃣ UPDATE existing record
#     if items:
#         item_id = items[0]["id"]
#         update_url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"

#         res = requests.patch(update_url, headers=headers, json=fields)
#         if not res.ok:
#             return jsonify({"error": "Update failed", "details": res.text}), 500

#         return jsonify({"status": "updated"})

#     # 3️⃣ CREATE new record
#     create_url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
#     res = requests.post(create_url, headers=headers, json={"fields": fields})

#     if not res.ok:
#         return jsonify({"error": "Create failed", "details": res.text}), 500

#     return jsonify({"status": "created"})

# # --------------------------------------------------
# # STEP 5: Get Document Libraries
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries", methods=["GET"])
# def get_libraries(site_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/drives"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     return [
#         d for d in res.json().get("value", [])
#         if d.get("driveType") == "documentLibrary"
#     ]

# # --------------------------------------------------
# # STEP 6: Upload Document
# # --------------------------------------------------
# @app.route("/sharepoint/<site_id>/libraries/<drive_id>/upload", methods=["POST"])
# def upload_document(site_id, drive_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     if "file" not in request.files:
#         return jsonify({"error": "No file provided"}), 400

#     file = request.files["file"]

#     upload_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{file.filename}:/content"
#     upload_headers = {
#         "Authorization": headers["Authorization"],
#         "Content-Type": file.content_type or "application/octet-stream"
#     }

#     res = requests.put(upload_url, headers=upload_headers, data=file.read())
#     if not res.ok:
#         return jsonify({"error": "Upload failed", "details": res.text}), 500

#     return jsonify(res.json())

# # --------------------------------------------------
# # STEP 7: Get Users (People Picker)
# # --------------------------------------------------
# @app.route("/graph/users", methods=["GET"])
# def graph_get_users():
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/users?$select=displayName,mail,userPrincipalName"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

#     users = []
#     for u in res.json().get("value", []):
#         email = u.get("mail") or u.get("userPrincipalName")
#         if not email:
#             continue

#         users.append({
#             "displayName": u["displayName"],
#             "email": email,
#             "claims": f"i:0#.f|membership|{email}"
#         })

#     return jsonify(users)

# # --------------------------------------------------
# # HEALTH
# # --------------------------------------------------
# @app.route("/health", methods=["GET"])
# def health():
#     return jsonify({"status": "SharePoint Profile API running"})

# # --------------------------------------------------
# # RUN
# # --------------------------------------------------
# if __name__ == "__main__":
#     print("🚀 Backend running on http://localhost:5050")
#     app.run(host="127.0.0.1", port=5050, debug=False)
















































































































import requests
from flask import Flask, jsonify, request
from flask_cors import CORS  # <--- NEW IMPORT

app = Flask(__name__)
CORS(app)  # <--- THIS ENABLES THE CONNECTION

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

HOSTNAME = "aoscaustralia.sharepoint.com"
SITE_PATH = "/sites/CPA"

# --------------------------------------------------
# Helper: Get Delegated User Token
# --------------------------------------------------
def get_headers():
    auth = request.headers.get("Authorization")
    if not auth or not auth.startswith("Bearer "):
        return None

    return {
        "Authorization": auth,
        "Content-Type": "application/json"
    }

# --------------------------------------------------
# STEP 1: Resolve SharePoint Site
# --------------------------------------------------
@app.route("/sharepoint/site", methods=["GET"])
def get_site():
    headers = get_headers()
    if not headers:
        return jsonify({"error": "Missing Authorization header"}), 401

    url = f"{GRAPH_BASE}/sites/{HOSTNAME}:{SITE_PATH}"
    res = requests.get(url, headers=headers)

    if not res.ok:
        return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

    site = res.json()
    return jsonify({
        "site_id": site["id"],
        "displayName": site["displayName"],
        "webUrl": site["webUrl"]
    })

# --------------------------------------------------
# STEP 2: Get Lists
# --------------------------------------------------
@app.route("/sharepoint/<site_id>/lists", methods=["GET"])
def get_lists(site_id):
    headers = get_headers()
    if not headers:
        return jsonify({"error": "Missing Authorization header"}), 401

    url = f"{GRAPH_BASE}/sites/{site_id}/lists"
    res = requests.get(url, headers=headers)

    if not res.ok:
        return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

    lists = [{
        "list_id": lst["id"],
        "displayName": lst.get("displayName"),
        "template": lst.get("list", {}).get("template")
    } for lst in res.json().get("value", [])]

    return jsonify(lists)

# --------------------------------------------------
# STEP 3: Get List Items
# --------------------------------------------------
@app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["GET"])
def get_list_items(site_id, list_id):
    headers = get_headers()
    if not headers:
        return jsonify({"error": "Missing Authorization header"}), 401

    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?expand=fields"
    res = requests.get(url, headers=headers)

    if not res.ok:
        return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

    return jsonify(res.json().get("value", []))

# --------------------------------------------------
# STEP 4: Create List Item (POST) – FULLY FIXED
# --------------------------------------------------
# @app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["POST"])
# def create_list_item(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     body = request.get_json()
#     if not body or "fields" not in body:
#         return jsonify({"error": "fields object required"}), 400

#     fields = body["fields"]

#     # 🚫 Remove system / calculated fields
#     for f in ["PercentComplete", "LastUpdated", "Created", "Modified"]:
#         fields.pop(f, None)

#     # ✅ AssignedStaff – SINGLE select only
#     if "AssignedStaff" in fields:
#         staff = fields["AssignedStaff"]

#         if not isinstance(staff, dict) or "claims" not in staff:
#             return jsonify({
#                 "error": "AssignedStaff must be a single claims object"
#             }), 400

#     print("POSTING FIELDS TO GRAPH:", fields)

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
#     res = requests.post(url, headers=headers, json={"fields": fields})

#     if not res.ok:
#         return jsonify({
#             "error": "Graph API error",
#             "details": res.text
#         }), res.status_code

#     return jsonify(res.json()), 201

# @app.route("/sharepoint/<site_id>/lists/<list_id>/columns", methods=["GET"])
# def get_list_columns(site_id, list_id):
#     headers = get_headers()
#     if not headers:
#         return jsonify({"error": "Missing Authorization header"}), 401

#     url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/columns"
#     res = requests.get(url, headers=headers)

#     if not res.ok:
#         return jsonify({
#             "error": "Graph API error",
#             "details": res.text
#         }), res.status_code

#     return jsonify(res.json())


# --------------------------------------------------
# STEP 4: UPSERT ITEM (With "Prefer" Header Fix)
# --------------------------------------------------
@app.route("/sharepoint/<site_id>/lists/<list_id>/items", methods=["POST"])
def upsert_list_item(site_id, list_id):
    headers = get_headers()
    if not headers:
        return jsonify({"error": "Missing Authorization header"}), 401

    body = request.get_json()
    if not body or "fields" not in body:
        return jsonify({"error": "fields object required"}), 400

    fields = body["fields"]
    email = fields.get("EmailAddress")
    
    if not email:
        return jsonify({"error": "EmailAddress is required to check for existing profile"}), 400

    # ✅ THE FIX: Add this special header to allow searching by Email
    search_headers = headers.copy()
    search_headers["Prefer"] = "HonorNonIndexedQueriesWarningMayFailRandomly"

    # Search for existing user by Email
    search_url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?expand=fields&$filter=fields/EmailAddress eq '{email}'"
    
    # Note: We use 'search_headers' here instead of just 'headers'
    search_res = requests.get(search_url, headers=search_headers)
    
    if not search_res.ok:
        return jsonify({"error": "Failed to search list", "details": search_res.text}), 500

    search_data = search_res.json()
    existing_items = search_data.get("value", [])

    # Cleanup system fields
    for f in ["PercentComplete", "LastUpdated", "Created", "Modified", "ID", "Author", "Editor"]:
        fields.pop(f, None)

    if len(existing_items) > 0:
        # --- UPDATE (PATCH) ---
        item_id = existing_items[0]["id"]
        print(f"🔄 Found existing profile (ID: {item_id}). Updating...")
        update_url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
        update_res = requests.patch(update_url, headers=headers, json=fields)
        if not update_res.ok:
            return jsonify({"error": "Failed to update item", "details": update_res.text}), 500
        return jsonify({"status": "Updated", "id": item_id, "fields": fields}), 200

    else:
        # --- CREATE (POST) ---
        print("🆕 No profile found. Creating new...")
        create_url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items"
        create_res = requests.post(create_url, headers=headers, json={"fields": fields})
        if not create_res.ok:
            return jsonify({"error": "Failed to create item", "details": create_res.text}), 500
        return jsonify(create_res.json()), 201
    




# --------------------------------------------------
# STEP 5: Get Document Libraries
# --------------------------------------------------
@app.route("/sharepoint/<site_id>/libraries", methods=["GET"])
def get_libraries(site_id):
    headers = get_headers()
    if not headers:
        return jsonify({"error": "Missing Authorization header"}), 401

    url = f"{GRAPH_BASE}/sites/{site_id}/drives"
    res = requests.get(url, headers=headers)

    if not res.ok:
        return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

    libraries = [
        d for d in res.json().get("value", [])
        if d.get("driveType") == "documentLibrary"
    ]

    return jsonify(libraries)

# --------------------------------------------------
# STEP 6: Get Documents
# --------------------------------------------------
@app.route("/sharepoint/<site_id>/libraries/<drive_id>/documents", methods=["GET"])
def get_documents(site_id, drive_id):
    headers = get_headers()
    if not headers:
        return jsonify({"error": "Missing Authorization header"}), 401

    url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"
    res = requests.get(url, headers=headers)

    if not res.ok:
        return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

    return jsonify(res.json().get("value", []))

# --------------------------------------------------
# STEP 7: Upload Document
# --------------------------------------------------
@app.route("/sharepoint/<site_id>/libraries/<drive_id>/upload", methods=["POST"])
def upload_document(site_id, drive_id):
    headers = get_headers()
    if not headers:
        return jsonify({"error": "Missing Authorization header"}), 401

    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]

    upload_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{file.filename}:/content"

    upload_headers = {
        "Authorization": headers["Authorization"],
        "Content-Type": file.content_type or "application/octet-stream"
    }

    res = requests.put(upload_url, headers=upload_headers, data=file.read())

    if not res.ok:
        return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

    return jsonify(res.json())

# --------------------------------------------------
# STEP 8: Get Users (People Picker)
# --------------------------------------------------
@app.route("/graph/users", methods=["GET"])
def graph_get_users():
    headers = get_headers()
    if not headers:
        return jsonify({"error": "Missing Authorization header"}), 401

    url = f"{GRAPH_BASE}/users?$select=displayName,mail,userPrincipalName"
    res = requests.get(url, headers=headers)

    if not res.ok:
        return jsonify({"error": "Graph API error", "details": res.text}), res.status_code

    users = []
    for u in res.json().get("value", []):
        email = u.get("mail") or u.get("userPrincipalName")
        if not email:
            continue

        users.append({
            "displayName": u["displayName"],
            "email": email,
            "claims": f"i:0#.f|membership|{email}"
        })

    return jsonify(users)

# --------------------------------------------------
# HEALTH
# --------------------------------------------------
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "Delegated SharePoint API running"})

# --------------------------------------------------
# RUN
# --------------------------------------------------
if __name__ == "__main__":
    print("🚀 Backend running on http://localhost:5050")
    app.run(host="127.0.0.1", port=5050, debug=False)
