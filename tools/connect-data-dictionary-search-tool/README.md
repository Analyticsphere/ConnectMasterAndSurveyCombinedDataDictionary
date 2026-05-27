# Connect Data Dictionary Search Tool

## Overview

The Connect Data Dictionary Search Tool enables investigators, analysts, and researchers to query the Connect Data Dictionary quickly and accurately without manually downloading the Excel file from GitHub or searching it by hand.

The tool is dynamic: each query fetches the latest available version of the data dictionary directly from GitHub, so results always reflect the most current content.

It is built for research and analysis workflows that require rapid access to concept IDs, variable labels, question text, response options, GCP table locations, and metadata needed to support data dictionary review.

> **Note:** This tool is a read-only lookup interface. It does not generate, edit, or modify the data dictionary in any way.

---

## What You Can Search For

The tool accepts natural-language queries. Examples:

- *What is the CID for self-reported sex?*
- *Look up CID 905787778.*
- *Find variables related to HIPAA revocation.*
- *Show the variable name, question text, response options, and GCP table location for a CID.*
- *Show whether a variable is marked as deprecated, new, or revised.*

Standard lookup results include the following fields, where available:

| Field | Description |
|---|---|
| Concept ID (CID) | Unique identifier for the variable |
| Variable Name | Programmatic variable name |
| Question Text | Current question text as it appears in the survey |
| Variable Label | Short descriptive label |
| Response Options | Coded response values, when applicable |
| Required Flag | Whether the variable is required |
| PII Flag | Whether the variable contains personally identifiable information |
| Status | Deprecated, New, or Revised designation |
| Production Date | Date the change was pushed to production, if recorded |
| GCP Location | Google Cloud Platform document/table location |
| Source Row | Row reference in the data dictionary workbook |

---

## Privacy and Data Handling

The data dictionary itself contains variable metadata, not participant-level data. However, because this tool supports analysis workflows that may involve sensitive study information, the following guidelines apply:

- Do **not** upload screenshots, query outputs, Box files, SharePoint files, or any participant-level information to a public repository.
- The downloaded data dictionary workbook, generated search index files, and your personal reference library are intended for local use only and should not be committed to version control.

---

## Public GitHub Upload Reference

### ✅ Safe to Upload

- `connect_dictionary.py`
- `README.md`
- `.gitignore`
- `data/.gitkeep`
- `library/connect_reference_library.example.json`

### ❌ Do Not Upload

- `data/MasterSurveyComb_latest.xlsx`
- `data/connect_dictionary_index.json`
- `data/workbook_metadata.json`
- `library/connect_reference_library.json`
- Screenshots or exported reports
- Files copied from Box or SharePoint unless explicitly approved for public release
- Query results or any participant-level data

---

## Getting the Tool

Before using the tool, the user needs a local copy of the GitHub folder that contains `connect_dictionary.py`.

### Option 1: Get the Tool From Terminal

1. Open the Mac Terminal app.
2. Go to the folder where the repository should be saved. Example:

```bash
cd ~/Documents
```

3. Copy and paste this command:

```bash
git clone https://github.com/Analyticsphere/ConnectMasterAndSurveyCombinedDataDictionary.git
```

4. Go to the tool folder:

```bash
cd ConnectMasterAndSurveyCombinedDataDictionary/tools/connect-data-dictionary-search-tool
```

5. Confirm the tool file is present:

```bash
ls
```

The folder should include `connect_dictionary.py`, `README.md`, `data`, and `library`.

### Option 2: Get the Tool From the GitHub Page

1. Open the GitHub tool page:
   <https://github.com/Analyticsphere/ConnectMasterAndSurveyCombinedDataDictionary/tree/main/tools/connect-data-dictionary-search-tool>
2. Click `Code`.
3. Click `Download ZIP`.
4. Unzip the downloaded file.
5. Open this folder inside the unzipped repository:

```text
ConnectMasterAndSurveyCombinedDataDictionary-main/tools/connect-data-dictionary-search-tool
```

The user does not need to manually download the Excel data dictionary. The tool downloads and indexes the workbook when it runs.

---

## Instructions for Use

The Connect Data Dictionary Search Tool can be used in several ways, depending on how someone prefers to work. The main file is `connect_dictionary.py`.

### Use Through Mac Terminal

1. Open the Mac Terminal app.
2. Go to the tool folder. If the repository was saved in `Documents` using `git clone`, use:

```bash
cd ~/Documents/ConnectMasterAndSurveyCombinedDataDictionary/tools/connect-data-dictionary-search-tool
```

3. Confirm Python 3 is available:

```bash
python3 --version
```

4. Refresh the data dictionary:

```bash
python3 connect_dictionary.py refresh
```

5. Run a lookup by CID:

```bash
python3 connect_dictionary.py cid 905787778
```

6. Or run a keyword search:

```bash
python3 connect_dictionary.py search "HIPAA revocation"
```

7. Or search for the best CID match from plain language:

```bash
python3 connect_dictionary.py cid "self reported sex"
```

### Use Through VSCode

1. Open VSCode.
2. Go to `File > Open Folder`.
3. Select the `connect-data-dictionary-search-tool` folder.
4. Open the VSCode terminal:

```text
Terminal > New Terminal
```

5. Refresh the data dictionary:

```bash
python3 connect_dictionary.py refresh
```

6. Search for a term:

```bash
python3 connect_dictionary.py search "verification date"
```

7. Look up a CID:

```bash
python3 connect_dictionary.py cid 905787778
```

8. Look up the best CID match from plain language:

```bash
python3 connect_dictionary.py cid "self reported sex"
```

### Use Through Codex

1. Open Codex.
2. Open the folder that contains `connect_dictionary.py`.
3. Ask Codex to use the local tool. Example:

```text
Use connect_dictionary.py to look up CID 905787778.
Refresh the workbook first.
Return the matched label/question, CID/state attribute, source context, row location, PII flag, GCP Document/Table, Deprecated/New/Revised status, and production date if available.
```

4. If Codex asks for permission to access GitHub, approve it so the tool can refresh the workbook.

Codex can run the search tool, pull the latest version of the Connect Data Dictionary from GitHub, and return the relevant information in a readable format.

### Refresh the Data Dictionary

The tool is designed to pull the latest available version of the Connect Data Dictionary before lookup searches. A user can also refresh the local copy manually:

```bash
python3 connect_dictionary.py refresh
```

### Create a Local Reference Library

Users can save frequently used lookup results to a local reference library. This is useful for variables, CIDs, or response options that are referenced often.

Save a result:

```bash
python3 connect_dictionary.py remember "self reported sex"
```

View saved references:

```bash
python3 connect_dictionary.py library
```

Show one saved reference:

```bash
python3 connect_dictionary.py show "self reported sex"
```

Remove a saved reference:

```bash
python3 connect_dictionary.py forget "self reported sex"
```

The local reference library is intended for personal use and should not be uploaded to GitHub.
