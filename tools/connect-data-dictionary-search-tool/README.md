Connect Data Dictionary Search Tool
Overview
The Connect Data Dictionary Search Tool enables investigators, analysts, and researchers to query the Connect Data Dictionary quickly and accurately — without manually downloading the Excel file from GitHub or searching it by hand.

The tool is dynamic: each query fetches the latest available version of the data dictionary directly from GitHub, so results always reflect the most current content.

It is built for research and analysis workflows that require rapid access to concept IDs, variable labels, question text, response options, GCP table locations, and metadata needed to support data dictionary change requests.

Note: This tool is a read-only lookup interface. It does not generate, edit, or modify the data dictionary in any way.

What You Can Search For
The tool accepts natural-language queries. Examples:

What is the CID for self-reported sex?
Look up CID 905787778.
Find variables related to HIPAA revocation.
Show the variable name, question text, response options, and GCP table location for a CID.
Show whether a variable is marked as deprecated, new, or revised.
Standard lookup results include the following fields, where available:

Field	Description
Concept ID (CID)	Unique identifier for the variable
Variable Name	Programmatic variable name
Question Text	Current question text as it appears in the survey
Variable Label	Short descriptive label
Response Options	Coded response values, when applicable
Required Flag	Whether the variable is required
PII Flag	Whether the variable contains personally identifiable information
Status	Deprecated, New, or Revised designation
Production Date	Date the change was pushed to production, if recorded
GCP Location	Google Cloud Platform document/table location
Source Row	Row reference in the data dictionary workbook
Privacy and Data Handling
The data dictionary itself contains variable metadata — not participant-level data. However, because this tool supports analysis workflows that may involve sensitive study information, the following guidelines apply:

Do not upload screenshots, query outputs, Box files, SharePoint files, or any participant-level information to a public repository.
The downloaded data dictionary workbook, generated search index files, and your personal reference library are intended for local use only and should not be committed to version control.
Public GitHub Upload Reference
✅ Safe to Upload
connect_dictionary.py
README.md
.gitignore
data/.gitkeep
library/connect_reference_library.example.json
❌ Do Not Upload
data/MasterSurveyComb_latest.xlsx
data/connect_dictionary_index.json
data/workbook_metadata.json
library/connect_reference_library.json
Screenshots or exported reports
Files copied from Box or SharePoint unless explicitly approved for public release
Query results or any participant-level data
Instructions for Use
This application is designed to run through Codex, though it can also be run locally on your computer. The main entry point is connect_dictionary.py.

Run the tool:

python3 connect_dictionary.py
Example with inline query:

python3 connect_dictionary.py cid "self reported sex"
