# Gmail Sender

It sends emails to a list of recipients based on :
 - Google Spreadsheet API for template & recipients
 - Google Drive API for downloading attachments
 - Google Gmail API to send emails

### Description

Defines a spreadsheet with 2 tabs : `Email` Tab & `Participants` Tab

`Email` tab contains two columns :
 - column *A* contains the keys (`Subject`, `Message`, `Attachment1`, `Attachment2` ...)
 - column *B* contains the email templates or attachments

Here are expected list of keys :
 - `Subject` which contains the email subject
 - `Message` which contains the email body
 - `Attachment` which contains an attachment sent with the email

Subject & Message are expected to store email template.
You can use `{Participants.ColumnName}` format to make each email personal.

Attachments are expected to be stored in Google Drive.
Format expected in attachment value is : `{file_name} / {file_id}` with
 - `file_name` is the one displayed in email attachment
 - `file_id` is the google drive file identifier

### How To use

Get Google certificates :
 - using https://developers.google.com/workspace/guides/create-credentials
 - requesting following scopes access :
  - https://www.googleapis.com/auth/spreadsheets
  - https://www.googleapis.com/auth/drive.readonly
  - https://www.googleapis.com/auth/gmail.send

Install poetry : https://python-poetry.org/docs/#installation

Execute program (but be careful!) :

```bash
python main.py -h
usage: email-sender [-h] [--credentials-path CREDENTIALS_PATH]
                    spreadsheet_id sender_email

send emails to list of participants

positional arguments:
  spreadsheet_id
  sender_email

options:
  -h, --help            show this help message and exit
  --credentials-path CREDENTIALS_PATH
```
