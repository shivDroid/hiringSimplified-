import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime
import time
from typing import Dict, List

# Google Sheets and Gmail API scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://www.googleapis.com/auth/gmail.send']

# Configure these values
SPREADSHEET_ID = 'https://docs.google.com/spreadsheets/d/1up6rKgZLHXxqLyl6QyTXmPUZcFoelGrUaDAGxcbbZd4'
RANGE_NAME = 'Sheet1!A:E'  # Adjust based on your sheet
SENDER_EMAIL = "abc@gmail.com"

class ApplicantTracker:
    def __init__(self):
        self.creds = self._get_credentials()
        self.sheets_service = build('sheets', 'v4', credentials=self.creds)
        self.last_processed = {}  # Store last processed status for each applicant
        
    def _get_credentials(self):
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
                
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        return creds

    def _get_sheet_data(self) -> List[Dict]:
        result = self.sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
        rows = result.get('values', [])
        
        if not rows:
            return []
            
        headers = rows[0]
        return [dict(zip(headers, row)) for row in rows[1:]]

    def _send_email(self, to_email: str, subject: str, body: str):
        service = build('gmail', 'v1', credentials=self.creds)
        
        message = MIMEMultipart()
        message['to'] = to_email
        message['subject'] = subject
        message['from'] = SENDER_EMAIL
        
        msg = MIMEText(body, 'html')
        message.attach(msg)
        
        raw_message = {'raw': base64.urlsafe_b64encode(
            message.as_bytes()).decode('utf-8')}
        
        try:
            service.users().messages().send(
                userId='me', body=raw_message).execute()
            print(f"Email sent to {to_email}")
        except Exception as e:
            print(f"Error sending email to {to_email}: {str(e)}")

    def _get_email_template(self, status: str) -> tuple:
        templates = {
            "NEW": (
                "Application Received - Thank You",
                """
                <p>Dear {name},</p>
                <p>Thank you for submitting your application. We have received it and will review it shortly.</p>
                <p>We appreciate your interest in joining our team.</p>
                <p>Best regards,<br>Recruitment Team</p>
                """
            ),
            "REJECTED": (
                "Application Status Update",
                """
                <p>Dear {name},</p>
                <p>Thank you for your interest in joining our team. After careful consideration, we regret to inform you 
                that your experience does not align with our current requirements.</p>
                <p>We wish you the best in your future endeavors.</p>
                <p>Best regards,<br>Recruitment Team</p>
                """
            ),
            "FOR_INTERVIEW": (
                "Interview Invitation",
                """
                <p>Dear {name},</p>
                <p>We are pleased to inform you that you have been selected for an interview.</p>
                <p>Our interview process consists of two stages:</p>
                <ol>
                    <li>Initial Technical Assessment</li>
                    <li>Team Culture Fit Interview</li>
                </ol>
                <p>We will contact you shortly to schedule the first interview.</p>
                <p>Best regards,<br>Recruitment Team</p>
                """
            )
        }
        return templates.get(status, (None, None))

    def process_applications(self):
        applications = self._get_sheet_data()
        
        for application in applications:
            applicant_id = application.get('ID')
            current_status = application.get('Status')
            email = application.get('Email')
            name = application.get('Name')
            
            # Skip if no change in status
            if (applicant_id in self.last_processed and 
                self.last_processed[applicant_id] == current_status):
                continue
                
            subject, template = self._get_email_template(current_status)
            if subject and template and email:
                body = template.format(name=name)
                self._send_email(email, subject, body)
                
            self.last_processed[applicant_id] = current_status

    def run_daily(self):
        while True:
            now = datetime.datetime.now()
            if now.hour == 17:  # 5 PM
                print(f"Running process at {now}")
                self.process_applications()
                # Wait until next hour before checking again
                time.sleep(3600)
            else:
                # Check every 5 minutes
                time.sleep(300)

if __name__ == "__main__":
    tracker = ApplicantTracker()
    tracker.run_daily()