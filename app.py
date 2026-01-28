from flask import Flask, render_template, request, session, send_file
from pywhatkit import sendwhats_image, sendwhatmsg_instantly
import csv
import io
import secrets
import pandas as pd
import time
import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# Use static/Images/ for saving uploaded image attachments
IMAGE_UPLOAD_FOLDER = os.path.join(os.getcwd(), "static", "Images")
os.makedirs(IMAGE_UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def trial():
    return render_template('index.html')

@app.route('/upload_csv', methods=['POST'])
def upload_csv():
    csv_file = request.files['csv_file']

    if csv_file.filename == '':
        return 'No selected file'

    if csv_file:
        stream = io.StringIO(csv_file.stream.read().decode("UTF8"), newline=None)
        csv_input = csv.reader(stream)
        session['csv_data'] = list(csv_input)
        return 'CSV file uploaded successfully'
    
    return 'Invalid request'

@app.route('/send_message', methods=['POST'])
def send_message():
    message = request.form['message']
    attachment = request.files['attachment']
    
    csv_data = session.get('csv_data', [])
    if not csv_data:
        return "No CSV data found. Please upload a CSV first."

    # Find WhatsApp number column
    whatsapp_index = None
    for idx, header in enumerate(csv_data[0]):
        if 'whatsapp_numbers' in header.lower():
            whatsapp_index = idx
            break

    if whatsapp_index is None:
        return "No 'whatsapp_numbers' column found in CSV."

    whatsapp_numbers = [row[whatsapp_index].strip() for row in csv_data[1:]]

    for number in whatsapp_numbers:
        try:
            full_number = f"+91{number}"
            if attachment and attachment.filename:
                saved_path = os.path.join(IMAGE_UPLOAD_FOLDER, attachment.filename)
                attachment.save(saved_path)
                sendwhats_image(full_number, saved_path, caption=message)
                time.sleep(15)
            else:
                sendwhatmsg_instantly(full_number, message)
                time.sleep(10)

            print(f"Message sent to {full_number}")
        except Exception as e:
            print(f"Error sending to {number}: {e}")

    return "Messages sent successfully!"

# Google Sheets config
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SAMPLE_SPREADSHEET_ID = "1fkFI4_uaildO1QeogRPZcMExHQ5rB3bDAJyD6XuNraw"

@app.route('/main', methods=['GET'])
def main():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        def get_sheet_range(service):
            spreadsheet = service.spreadsheets().get(
                spreadsheetId=SAMPLE_SPREADSHEET_ID).execute()
            sheet_title = spreadsheet['sheets'][0]['properties']['title']
            return f"{sheet_title}!A1:ZZ"

        SAMPLE_RANGE_NAME = get_sheet_range(service)

        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=SAMPLE_RANGE_NAME
        ).execute()

        values = result.get("values", [])
        if not values:
            return "No data found."

        def filter_interested(values):
            headers = values[0]
            interested_col_index = headers.index('Are you interested in our college ?')
            interested_entries = [
                entry for entry in values[1:]
                if len(entry) > interested_col_index and entry[interested_col_index] == 'Yes'
            ]
            return [headers] + interested_entries

        df = pd.DataFrame(filter_interested(values))
        df.to_csv('output_file.csv', index=False, header=False)

        if not os.path.exists('output_file.csv'):
            return 'CSV data not found'

        return send_file('output_file.csv', as_attachment=True)

    except HttpError as err:
        return f"Google Sheets API error: {err}"

@app.route("/new")
def testing():
    return render_template("new.html")

if __name__ == '__main__':
    app.run(debug=True)
