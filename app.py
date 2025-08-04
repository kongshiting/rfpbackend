from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import gspread
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import time
import traceback
from googleapiclient.errors import HttpError

app = Flask(__name__)
CORS(app)

# Google API Setup
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive"
]

SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")
TEMPLATE_DOC_ID = os.getenv("TEMPLATE_DOC_ID")

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(creds)
sheet = client.open_by_key(SPREADSHEET_ID).sheet1

docs_service = build("docs", "v1", credentials=creds)
drive_service = build("drive", "v3", credentials=creds)

# Category-specific field mappings
CATEGORY_FIELDS = {
    "Meal": {
        "fields": ["participantList", "mealPurpose"],
        "doc_placeholders": {
            "{{ParticipantList}}": "participantList",
            "{{MealPurpose}}": "mealPurpose"
        }
    },
    "Transport": {
        "fields": ["transportType", "departure", "destination"],
        "doc_placeholders": {
            "{{TransportType}}": "transportType",
            "{{Departure}}": "departure",
            "{{Destination}}": "destination"
        }
    },
    "Printing": {
        "fields": ["printingDetails", "copiesCount"],
        "doc_placeholders": {
            "{{PrintingDetails}}": "printingDetails",
            "{{CopiesCount}}": "copiesCount"
        }
    }
}

@app.route("/submit", methods=["POST"])
def submit_form():
    try:
        # Parse form data
        print("Form data:", request.form)
        payee_name = request.form.get("payeeName")
        matric_no = request.form.get("matricNo")
        nus_net_id = request.form.get("nusNetID")
        category = request.form.get("category")
        event_name = request.form.get("eventName")
        committee = request.form.get("committee")
        expense_count = int(request.form.get("expenseCount", "0"))

        copy_metadata = {
            "name": f"RFP_{payee_name}_{int(time.time())}",
            "parents": [DRIVE_FOLDER_ID]
        }
        copied_doc = drive_service.files().copy(
            fileId=TEMPLATE_DOC_ID,
            body=copy_metadata,
            supportsAllDrives=True
        ).execute()
        doc_id = copied_doc.get("id")

        replace_text_in_doc(doc_id, {
            "{{Category}}": category,
            "{{EventName}}": event_name,
            "{{Committee}}": committee,
            "{{PayeeName}}": payee_name,
            "{{MatricNo}}": matric_no,
            "{{NUSNET}}": nus_net_id,
        })

        total_amount = 0
        expense_details = []
        for i in range(1, 5):
            amount_str = request.form.get(f"expense{i}amount", "0")
            try:
                amount = float(amount_str) if amount_str else 0
                total_amount += amount
            except ValueError:
                amount = 0

            expense_details.append({
                "receipt_no": request.form.get(f"expense{i}receiptno", ""),
                "description": request.form.get(f"expense{i}description", ""),
                "amount": amount_str,
                "purchase_type": request.form.get(f"expense{i}purchasetype", "")
            })

        # Prepare base replacements
        replacements = {
            "{{Category}}": category,
            "{{EventName}}": event_name,
            "{{Committee}}": committee,
            "{{PayeeName}}": payee_name,
            "{{MatricNo}}": matric_no,
            "{{NUSNET}}": nus_net_id,
            "{{TotalAmount}}": f"{total_amount:.2f}"
        }

        for i, expense in enumerate(expense_details, start=1):
            replacements.update({
                f"{{{{ReceiptNo{i}}}}}": expense["receipt_no"],
                f"{{{{Description{i}}}}}": expense["description"],
                f"{{{{Amount{i}}}}}": expense["amount"],
                f"{{{{PurchaseType{i}}}}}": expense["purchase_type"]
            })

        # Add category-specific fields
        if category in CATEGORY_FIELDS:
            for placeholder, field_name in CATEGORY_FIELDS[category]["doc_placeholders"].items():
                replacements[placeholder] = request.form.get(field_name, "")

        # Process all text replacements
        replace_text_in_doc(doc_id, replacements)

        file_urls = []
        for file_key in request.files:
            file = request.files[file_key]
            if file.filename:
                file_path = os.path.join("/tmp", file.filename)
                file.save(file_path)

                try:
                    uploaded_file = upload_to_drive(file_path)
                    file_url = uploaded_file.get("webViewLink")
                    file_urls.append(file_url)

                    if file.filename.lower().endswith(("png", "jpg", "jpeg")):
                        insert_image_to_doc(doc_id, uploaded_file.get("id"))
                    else:
                        append_text_to_doc(doc_id, f"\nAttachment: {file_url}")
                finally:
                    # Clean up the uploaded file
                    if os.path.exists(file_path):
                        os.remove(file_path)

        # Append to spreadsheet
        sheet_row = [
            payee_name,
            matric_no,
            category,
            event_name,
            committee,
            *[f"{exp['receipt_no']} | {exp['description']} | {exp['amount']} | {exp['purchase_type']}" for exp in expense_details],
            str(total_amount),
            ", ".join(file_urls)
        ]
        sheet.append_row(sheet_row)

        return jsonify({
            "message": "Form submitted successfully!",
            "doc_url": f"https://docs.google.com/document/d/{doc_id}"
        }), 200

    except Exception as e:
        print("===== ERROR TRACEBACK =====")
        traceback.print_exc()
        print("===========================")
        return jsonify({"error": str(e)}), 500


def replace_text_in_doc(doc_id, replacements):
    """Replaces placeholders in a Google Doc."""
    requests = []
    for key, val in replacements.items():
        if val:  # Only replace if there's a value
            requests.append({
                "replaceAllText": {
                    "containsText": {"text": key, "matchCase": True},
                    "replaceText": val
                }
            })
    
    if requests:
        docs_service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": requests}
        ).execute()

def append_text_to_doc(doc_id, text):
    """Appends text to a Google Doc."""
    # Get document to find the end index
    doc = docs_service.documents().get(documentId=doc_id).execute()
    end_index = doc['body']['content'][-1]['endIndex'] - 1
    
    requests = [{
        "insertText": {
            "location": {"index": end_index},
            "text": text + "\n"
        }
    }]
    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": requests}
    ).execute()

def insert_image_to_doc(doc_id, image_id):
    """Inserts an image into a Google Doc."""
    # Get document to find the end index
    doc = docs_service.documents().get(documentId=doc_id).execute()
    end_index = doc['body']['content'][-1]['endIndex'] - 1
    
    requests = [{
        "insertInlineImage": {
            "location": {"index": end_index},
            "uri": f"https://drive.google.com/uc?id={image_id}",
            "objectSize": {
                "height": {"magnitude": 300, "unit": "PT"},
                "width": {"magnitude": 400, "unit": "PT"}
            }
        }
    }]
    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": requests}
    ).execute()


def upload_to_drive(file_path):
    """Uploads a file to Google Drive and returns file metadata."""
    file_name = os.path.basename(file_path)
    mime_type = "application/pdf"
    
    media = MediaFileUpload(file_path, mimetype=mime_type, resumable=True)
    file_metadata = {
        "name": file_name,
        "parents": [DRIVE_FOLDER_ID]
    }
    
    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, webViewLink",
        supportsAllDrives=True
    ).execute()

    print(f"Uploaded file response: {uploaded_file}") 
    file_id = uploaded_file.get("id")
    print(f"Uploaded file ID: {file_id}")

    return uploaded_file

    def set_permission_with_retry(drive_service, file_id, max_retries=5, wait_seconds=5):
        for attempt in range(max_retries):
            try:
                drive_service.permissions().create(
                    fileId=file_id,
                    body={"type": "anyone", "role": "reader"},
                    fields="id",
                    supportsAllDrives=True
                ).execute()
                print("✅ Permission set.")
                return
            except HttpError as e:
                if e.resp.status == 404:
                    print(f"File not ready yet (attempt {attempt+1})")
                    time.sleep(wait_seconds)
                else:
                    print(f"Unexpected error: {e}")
                    raise
        raise Exception("File never became ready for permissions.")

@app.route("/category-fields/<category>", methods=["GET"])
def get_category_fields(category):
    """Returns the additional fields required for a specific category."""
    fields = CATEGORY_FIELDS.get(category, {}).get("fields", [])
    return jsonify({"fields": fields})

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(debug=True)


