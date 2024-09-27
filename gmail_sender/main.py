import argparse
import base64
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
import io
import mimetypes
import os.path
import re
import smtplib

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/gmail.send"
]

def get_credentials(credential_path) -> any:
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  creds = None

  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
        credential_path, SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  return creds

def get_participants(spreadsheet_id, sheet: any) -> list:
    result = (
        sheet.values()
        .get(spreadsheetId=spreadsheet_id, range="Participants!A:R")
        .execute()
    )
    values = result.get("values", [])

    if not values:
      print("No participant found")
      return

    keys = []
    participants = []

    for row in values:
      if len(keys) == 0:
        keys = ["Participants." + item for item in row]
        continue

      participants.append(dict(zip(keys, row)))

    return participants

def get_email(spreadsheet_id, sheet: any) -> dict:
    result = (
        sheet.values()
        .get(spreadsheetId=spreadsheet_id, range="Email!A:B")
        .execute()
    )
    values = result.get("values", [])

    if not values:
      print("No email found")
      return

    email = {}

    for row in values:
      key = row[0]
      value = row[1]
      email[key] = value

    return email

def get_attachments(creds, email):
  attachments = {}

  try:
    # create drive api client
    service = build("drive", "v3", credentials=creds)

    for key, value in email.items():
      if key.startswith("Attachment"):
        file_name, file_id = value.split("/")
        attachments[file_name.strip()] = file_id.strip()

    for file_name, file_id in attachments.items():
      request = service.files().get_media(fileId=file_id)
      file = io.BytesIO()
      downloader = MediaIoBaseDownload(file, request)
      done = False
      while done is False:
        status, done = downloader.next_chunk()
        print(f"Download {file_name}: {int(status.progress() * 100)}.")

      attachments[file_name] = file.getvalue()

  except HttpError as error:
    print(f"An error occurred: {error}")
    file = None

  return attachments

def generate_emails_to_send(email: dict, participants: list):
  emails_to_send = []

  for participant in participants:
    email_to_send = {
      "To": participant["Participants.Email payeur"],
    }

    for key, value in email.items():
      matches = re.finditer(r"{([^}]+)}", value, re.MULTILINE)

      for match in matches:
        to_replace = match.group(0)
        participant_key = match.group(1)

        if participant_key in participant:
          value = value.replace(to_replace, participant[participant_key])

      email_to_send[key] = value

    emails_to_send.append(email_to_send)

  return emails_to_send

def send_gmail_email(creds, sender_email, email, attachments: list):
  try:
    recipient = email["To"]

    # create gmail api client
    service = build("gmail", "v1", credentials=creds)

    message = EmailMessage()
    message["From"] = sender_email
    message["To"] = recipient
    message["Subject"] = email["Subject"]

    plain_text_content = email["Message"]
    html_content = "<html><body>" + plain_text_content.replace("\n", "<br>") +"</body></html>"

    message.set_content(plain_text_content, subtype="plain")
    message.add_alternative(html_content, subtype="html")

    # attachments
    for file_name, file_data in attachments.items():
      # guessing the MIME type
      type_subtype, _ = mimetypes.guess_type(file_name)
      maintype, subtype = type_subtype.split("/")
      message.add_attachment(
        file_data,
        maintype=maintype,
        subtype=subtype,
        filename=file_name,
      )

    # encoded message
    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    create_message = {"raw": encoded_message}
    # pylint: disable=E1101
    send_message = (
        service.users()
        .messages()
        .send(userId="me", body=create_message)
        .execute()
    )
    print(f'Email sent to {recipient}: {send_message["id"]}')
  except HttpError as error:
    print(f"An error occurred: {error}")
    send_message = None

  return send_message

def main():
  parser = argparse.ArgumentParser(
    prog="email-sender",
    description="send emails to list of participants"
  )
  parser.add_argument("spreadsheet_id")
  parser.add_argument("sender_email")
  parser.add_argument("--credentials-path", default="credentials.json")

  args = parser.parse_args()

  spreadsheet_id = args.spreadsheet_id
  if spreadsheet_id == "":
      parser.print_help()
      return

  creds = get_credentials(args.credentials_path)

  try:
    # Call the Sheets API
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    participants = get_participants(spreadsheet_id, sheet)
    email = get_email(spreadsheet_id, sheet)

    # Call the Drive API
    attachments = get_attachments(creds, email)

    emails_to_send = generate_emails_to_send(email, participants)

    # Call the Gmail API
    for email_to_send in emails_to_send:
      send_gmail_email(creds, args.sender_email, email_to_send, attachments)

  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()
