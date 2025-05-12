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

@app.route("/submit", methods=["POST"])
def submit_form():
    try:
        # Parse form data
        print("Form data:", request.form)
        payee_name = request.form.get("payeeName")
        matric_no = request.form.get("matricNo")
        category = request.form.get("category")
        event_name = request.form.get("eventName")
        committee = request.form.get("committee")
        expense_count = int(request.form.get("expenseCount", "0"))

        copy_metadata = {"name": f"RFP_{payee_name}_{int(time.time())}", "parents": [DRIVE_FOLDER_ID]}
        copied_doc = drive_service.files().copy(fileId=TEMPLATE_DOC_ID, body=copy_metadata, supportsAllDrives=True).execute()
        doc_id = copied_doc.get("id")

        replace_text_in_doc(doc_id, {
            "{{Category}}": category,
            "{{EventName}}": event_name,
            "{{Committee}}": committee,
            "{{PayeeName}}": payee_name,
            "{{MatricNo}}": matric_no,
        })

        receipt_no_1 = request.form.get("expense1receiptno", "").strip()
        description_1 = request.form.get("expense1description", "").strip()
        amount_1 = request.form.get("expense1amount", "").strip()
        purchase_type_1 = request.form.get("expense1purchasetype", "").strip()
        replace_text_in_doc(doc_id, {
            "{{ReceiptNo1}}": receipt_no_1,
            "{{Description1}}": description_1,
            "{{Amount1}}": amount_1,
            "{{PurchaseType1}}": purchase_type_1
        })

        receipt_no_2 = request.form.get("expense2receiptno", "").strip()
        description_2 = request.form.get("expense2description", "").strip()
        amount_2 = request.form.get("expense2amount", "").strip()
        purchase_type_2 = request.form.get("expense2purchasetype", "").strip()
        replace_text_in_doc(doc_id, {
            "{{ReceiptNo2}}": receipt_no_2,
            "{{Description2}}": description_2,
            "{{Amount2}}": amount_2,
            "{{PurchaseType2}}": purchase_type_2
        })

        receipt_no_3 = request.form.get("expense3receiptno", "").strip()
        description_3 = request.form.get("expense3description", "").strip()
        amount_3 = request.form.get("expense3amount", "").strip()
        purchase_type_3 = request.form.get("expense3purchasetype", "").strip()
        replace_text_in_doc(doc_id, {
            "{{ReceiptNo3}}": receipt_no_3,
            "{{Description3}}": description_3,
            "{{Amount3}}": amount_3,
            "{{PurchaseType3}}": purchase_type_3
        })

        receipt_no_4 = request.form.get("expense4receiptno", "").strip()
        description_4 = request.form.get("expense4description", "").strip()
        amount_4 = request.form.get("expense4amount", "").strip()
        purchase_type_4 = request.form.get("expense4purchasetype", "").strip()
        replace_text_in_doc(doc_id, {
            "{{ReceiptNo4}}": receipt_no_4,
            "{{Description4}}": description_4,
            "{{Amount4}}": amount_4,
            "{{PurchaseType4}}": purchase_type_4
        })

        receipt_no_5 = request.form.get("expense5receiptno", "").strip()
        description_5 = request.form.get("expense5description", "").strip()
        amount_5 = request.form.get("expense5amount", "").strip()
        purchase_type_5 = request.form.get("expense5purchasetype", "").strip()
        replace_text_in_doc(doc_id, {
            "{{ReceiptNo5}}": receipt_no_5,
            "{{Description5}}": description_5,
            "{{Amount5}}": amount_5,
            "{{PurchaseType5}}": purchase_type_5
        })

        file_url = ""
        if "file" in request.files:
            file = request.files["file"]
            if file.filename:
                file_path = os.path.join("uploads", file.filename)
                file.save(file_path)

                uploaded_file = upload_to_drive(file_path)
                file_url = uploaded_file.get("webViewLink")

                if file.filename.lower().endswith(("png", "jpg", "jpeg")):
                    insert_image_to_doc(doc_id, uploaded_file.get("id"))
                else:
                    append_text_to_doc(doc_id, f"\nAttached PDF: {file_url}")

        sheet.append_row([
            payee_name, matric_no, category, event_name, committee,
            *[f"{request.form.get(f'expense{i}receiptno')} | {request.form.get(f'expense{i}description')} | {request.form.get(f'expense{i}amount')} | {request.form.get(f'expense{i}purchasetype')}" for i in range(1, expense_count + 1)]
        ])

        return jsonify({"message": "Form submitted successfully!", "doc_url": f"https://docs.google.com/document/d/{doc_id}"}), 200

    except Exception as e:
        print("===== ERROR TRACEBACK =====")
        traceback.print_exc()
        print("===========================")
        return jsonify({"error": str(e)}), 500


def replace_text_in_doc(doc_id, replacements):
    """Replaces placeholders in a Google Doc."""
    requests = [{"replaceAllText": {"containsText": {"text": key, "matchCase": True}, "replaceText": val}} for key, val in replacements.items()]
    docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()


def append_text_to_doc(doc_id, text):
    """Appends text to a Google Doc."""
    requests = [{"insertText": {"location": {"index": 1}, "text": text}}]
    docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()


def insert_image_to_doc(doc_id, image_id):
    """Inserts an image into a Google Doc."""
    requests = [{"insertInlineImage": {"location": {"index": 1}, "uri": f"https://drive.google.com/uc?id={image_id}"}}]
    docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()


def upload_to_drive(file_path):
    """Uploads a file to Google Drive and returns file metadata."""
    media = MediaFileUpload(file_path, resumable=True)
    file_metadata = {"name": os.path.basename(file_path), "parents": [DRIVE_FOLDER_ID]}
    return drive_service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink").execute()

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(debug=True)
    

