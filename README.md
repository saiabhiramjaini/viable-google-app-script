
# Email Invoice Processing Script — Google Apps Script

This script parses incoming Gmail messages with PDF attachments or inline images related to trial documents, converts them to PDF if needed, performs OCR, parses the extracted text, saves them to Google Drive, and logs their details in a Google Sheet.


[Watch video on YouTube](https://www.youtube.com/watch?v=wHY16iZ1xpQ)


---

## Installation and Setup

### 1. Set up your Google Drive folder:

* Create a folder in Google Drive where processed files will be stored.
* **Copy its folder ID from the URL.**
  Your folder’s URL might look like this:

```
https://drive.google.com/drive/folders/<FOLDER_ID>
```

---

### 2. Set up your Google Sheet:

* Create a new Google Sheet.
* **Copy its sheet ID from the URL.**
  Your sheet’s URL might look like this:

```
https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit#gid=0
```


---

### 3. Update IDs in script:

Open **Apps Script** ([https://script.google.com/](https://script.google.com/)) related to your Google Workspace and **set the following IDs at the top of your script**:

```javascript
const SHEET_ID = "your-sheet-ID";
const FOLDER_ID = "your-drive-folder-ID";
```

---

## How to deploy and use

### 1. Open Google Apps Script:

* From your **Google Drive**, create or open **Google Apps Script**.

---

### 2. Copy and paste script:

* Remove the existing code.
* **Paste the script provided above in your script editor**.

---

### 3. Save and authorize:

* Save the script.
* The first time you run `processEmails()` or set up a trigger, **Google will ask you to authorize the script** with your Google account.
* **Grant the necessary permissions** (Gmail, Drive, Docs, and Sheets).

---

### 4. Create a time-based trigger:

To enable automated processing:

```javascript
function createTrigger() {
  ScriptApp.newTrigger("processEmails")
    .timeBased()
    .everyHours(3)
    .create();

  Logger.log("Created new trigger to run every 3 hours.");
}
```

Running `createTrigger()` once from your script editor will set up **a time-based trigger** that runs `processEmails()` **every 3 hours**.

---

## How the script works

* **Search for matching emails** with subject starting with "Viable: Trial Document".

* If matching messages are found:

  * parses PDF attachments
  * parses and converts inline images to PDF

* Performs **OCR** to extract text from documents.

* Saves parsed files in your Google Drive folder.

* Logs parsed data (Invoice Number, Amount, Date, Vendor, File URL) in your Google Sheet.

---

## File structure

* **Apps Script** — parses emails, converts files, parses text.
* **Google Drive** — stores processed files.
* **Google Sheets** — maintains a record with parsed details.

---
