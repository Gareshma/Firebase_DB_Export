# Firestore Sessions to Excel Export

This guide explains how to export your Firestore `sessions` collection into a single Excel workbook where:

- each **session document** becomes **one sheet**
- the **session fields** appear at the top of the sheet
- the **questions subcollection** appears below as a table

This is designed for a structure like:

- `/sessions/{sessionId}`
- `/sessions/{sessionId}/questions/{questionId}`

## What the export creates

The script creates one Excel file:

- `sessions_one_document_per_sheet_polished.xlsx`

Inside that workbook:

- **one sheet per session document**
- **top section:** session metadata such as `sessionId`, `cardNumber`, `grade`, `createdAt`, etc.
- **bottom section:** question rows from the `questions` subcollection

## Prerequisites

Make sure you have:

- Python 3.9 or newer
- access to your Firebase project
- a Firebase service account key JSON file

## Step 1: Download your Firebase service account key

In Firebase Console:

1. Open your project
2. Go to **Project settings**
3. Open the **Service accounts** tab
4. Click **Generate new private key**
5. Download the JSON file

Save that file in the same folder as your Python script and rename it to:

- `serviceAccountKey.json`

## Step 2: Install required packages

Open Command Prompt, PowerShell, or terminal in the project folder and run:

```bash
pip install firebase-admin pandas openpyxl
```

## Step 3: Create the Python script

Create a file named:

- `export_firestore_to_excel.py`

Paste the export script into that file.

## Step 4: Recommended folder structure

Your folder should look like this before running:

```text
your-folder/
|- export_firestore_to_excel.py
|- serviceAccountKey.json
```

## Step 5: Run the export

In the same folder, run:

```bash
python export_firestore_to_excel.py
```

If everything works, the script will generate:

```text
sessions_one_document_per_sheet_polished.xlsx
```

After running, your folder will look like:

```text
your-folder/
|- export_firestore_to_excel.py
|- serviceAccountKey.json
|- sessions_one_document_per_sheet_polished.xlsx
```

## How the Excel workbook is organized

For each session document:

- the sheet name is based on the `sessionId`
- invalid Excel sheet name characters are automatically cleaned
- if the sheet name is too long, it is shortened to fit Excel's 31-character limit

Inside each sheet:

### Session section
The first rows contain key-value pairs such as:

- `sessionId`
- `cardNumber`
- `grade`
- `createdAt`
- `status`
- `score`
- `participantId`
- `name`

### Questions section
Below the session section, a **Questions** heading appears, followed by a table containing the question documents.

Typical columns may include:

- `questionId`
- `questionNumber`
- `userAnswer`
- `correctAnswer`
- `correct`
- `speakerOn`
- `audioEnPlayedMs`
- `audioEsPlayedMs`
- `timeSpentSec`
- `timeSpentMin`

The exact columns depend on what is stored in each question document.

## Important fix: scrolling issue in Excel

If you open the workbook and it looks like you cannot scroll down properly, the cause is usually **freeze panes**.

### Quick fix inside Excel

1. Open the workbook
2. Go to **View**
3. Click **Freeze Panes**
4. Click **Unfreeze Panes**

### Better fix in the Python script

Remove or disable the line that freezes rows, for example:

```python
ws.freeze_panes = f"A{questions_data_start_row}"
```

Replace it with:

```python
ws.freeze_panes = None
```

This makes scrolling normal and is the recommended setup for this layout.

## Common issues and fixes

### 1) `FileNotFoundError: serviceAccountKey.json`
Cause:
- The service account file is missing or not in the same folder.

Fix:
- Make sure `serviceAccountKey.json` is in the same folder as the script.

### 2) `ModuleNotFoundError`
Cause:
- Required Python packages are not installed.

Fix:

```bash
pip install firebase-admin pandas openpyxl
```

### 3) Permission errors from Firebase
Cause:
- Wrong project key or restricted service account.

Fix:
- Regenerate the service account key from the correct Firebase project.
- Make sure the key belongs to the project that contains the `sessions` collection.

### 4) Excel opens but some sheets look wide
Cause:
- Question documents can have many fields.

Fix:
- This is normal. Use Excel horizontal scroll or filter columns.
- If needed, modify the script to export only selected columns.

## Security note

Your `serviceAccountKey.json` is sensitive.

Do not:

- upload it to GitHub
- share it publicly
- commit it into your repository

Add it to `.gitignore` if your folder is inside a git project.

Example:

```gitignore
serviceAccountKey.json
```

## Optional improvements you can add later

You can enhance the export script to:

- include the `events` subcollection
- export only selected fields
- sort questions by `questionNumber`
- color-code correct vs incorrect answers
- apply filters to the question table
- add a summary sheet across all sessions

## Recommended workflow

1. Keep the export script in a separate utility folder
2. Run the script whenever you need a fresh export
3. Open the generated workbook in Excel
4. If scrolling feels odd, unfreeze panes
5. Save a copy if you want to share the file without rerunning the script

## Summary

This export setup gives you a clean way to convert your Firestore data into Excel:

- one workbook
- one sheet per session
- session details at the top
- all related questions below

It works well for reviewing quiz/session data in a readable, shareable format.
