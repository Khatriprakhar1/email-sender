# Bulk Email Sender (Gmail + Excel)

## Setup (one time)
1. Install Python 3.9+
2. `pip install -r requirements.txt`
3. Turn on 2-Step Verification in your Google Account, then create an App Password:
   Google Account -> Security -> 2-Step Verification -> App passwords (16 characters).

## Run
    streamlit run app.py

A browser page opens. Enter Gmail + App Password (sidebar), upload the Excel file,
write subject/body using placeholders like {name} and {company}, add attachments,
send a TEST email to yourself, then click Send.

## Excel format
Any columns are fine. Example: Name | Company | Email
Placeholders use your column names, e.g. {Name}, {Company}, {City}.

## Output
You can download an Excel file with Status (Sent/Failed), Sent At and Error columns.
Upload that file again later to skip rows already Sent.
