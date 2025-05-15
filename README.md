# Wire Set Certificate Processing Script

![Last Updated](https://img.shields.io/badge/last%20updated-May%202024-informational)
![Status](https://img.shields.io/badge/status-production-brightgreen)
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.12-blue)
![OCR](https://img.shields.io/badge/OCR-Tesseract-lightgrey)
![Schedule](https://img.shields.io/badge/schedule-10min%20loop%205AM--5PM-blueviolet)


This script automates the association of **TC wire sets** with the **wire roll** they were made from, by reading each set's calibration certificate (PDF) and extracting the wire roll serial number using OCR. It interacts with the **Qualer SDK**, and **SharePoint** via the **Microsoft Graph API** to fetch and update data.

## Instructions for Pyrometry Inspectors

This script will do its best to associate wire sets with their wire rolls. However, this system is inherently fragile, as it relies on several underlying assumptions:

* All wire sets are present in [WireSetCerts.xlsx](https://jgiquality.sharepoint.com/:x:/s/JGI/Ed0TEK1rlx9EjiIk6tqYX7cBeNrpNLL4JyOxY30ts-qnZA?e=eeKWFF).
* Wire set certificates are saved to the Service Order on Qualer (⚠️ Not the Work Item).
* Wire set certificates are saved in `.pdf` format.
* Wire set certificates' filenames begin with their Asset Tag, as it appears in Qualer.
* Wire set certificates contain the wire roll serial number directly between "The above expendable wireset was made from wire roll " and the following ". ".
* The OCR process is able to correctly identify the serial number according to the previous assumption.
* Wire roll certificates are named according to the `{SerialNumber}.xls` convention (e.g., `011391A.xls`).
* Wire roll certificates are stored at [Pyro_Standards](https://jgiquality.sharepoint.com/sites/JGI/Shared%20Documents/Pyro/Pyro_Standards/).
* The SharePoint file (`WireSetCerts.xlsx`) is **not open or locked** during the upload attempt.
  ⚠️ If someone is editing the file in Excel (desktop or browser), SharePoint may return a **423 Locked** error and prevent the script from uploading. The script will silently retry on the next cycle.


> 💡 *Try to avoid keeping the file open in Excel for more than a few minutes during work hours.*

_If any one of these assumptions is violated, it is the responsibility of the Pyrometry department to **manually** update [WireSetCerts.xlsx](https://jgiquality.sharepoint.com/:x:/s/JGI/Ed0TEK1rlx9EjiIk6tqYX7cBeNrpNLL4JyOxY30ts-qnZA?e=eeKWFF) with the Wire Roll serial number whenever a new Wire Set is certified._

### Manual Update Instructions

1. Open [WireSetCerts.xlsx](https://jgiquality.sharepoint.com/:x:/s/JGI/Ed0TEK1rlx9EjiIk6tqYX7cBeNrpNLL4JyOxY30ts-qnZA?e=eeKWFF) in Excel.
2. Locate the row corresponding to the wire set.
3. Enter the correct wire roll serial number in the appropriate column.
4. Save the file and ensure it is uploaded back to SharePoint.

> **Note**: If a wire roll certificate is missing or misnamed, contact the responsible team to correct the issue. Ensure all naming conventions are followed to avoid future errors.

---

# IT Information

## 🚀 Features

* 🧠 **Smart Lookups**: Pulls the latest Qualer service record for each wire set and locates the matching certificate document.
* 📟 **OCR Extraction**: Converts PDFs to images and uses Tesseract to find the wire roll serial number within the certificate.
* 📊 **Excel Updating**: Updates the SharePoint-hosted `WireSetCerts.xlsx` file with extracted data.
* ☁️ **Integrated with SharePoint & Graph API**: Downloads the Excel file and uploads it back only if changes are detected.
* 🔐 **Secure Token Handling**: Pulls the Qualer API token from a protected `apikey.txt` file in SharePoint.
* 🔀 **Scheduled Execution**: Intended to run every 10 minutes from **5:00 AM to 5:00 PM**, Monday through Saturday.
* 🩵 **Detailed Logging**: Logs all steps and errors with timestamps using Python's `logging` module.

---

## 🧱 Prerequisites

* Python 3.8+
* [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (installed locally).
* Access to:

  * Microsoft Azure AD App credentials (client ID, secret, tenant ID).
  * Qualer API.
  * JGI's SharePoint site.
  * Host Server (Currently the **QualerCalSync** VM on `JGI-HV-11`)

---

## 📦 Installation

```bash
git clone https://github.com/your-org/tc-wires.git
cd tc-wires
pip install -r requirements.txt
```

---

## ⚙️ Configuration

Create a `.env` file in the root directory with:

```env
AZURE_CLIENT_ID=f0fc42a5-125f-4eed-ae6c-5b58ab9c0971
AZURE_CLIENT_SECRET=your-app-secret
AZURE_TENANT_ID=9def3ae4-854a-4465-952c-5693835965d9
SHAREPOINT_DRIVE_ID=b!34PQK-JF0EmH57ieExSqveCp2B5j30NMsNTGcMEXae_5x8SnfJhdR6JqUh5dD03F
SHAREPOINT_SITE_ID=jgiquality.sharepoint.com,b8d7ad55-622f-41e1-9140-35b87b4616f9,160cda33-41a0-4b31-8ebf-11196986b3e3
```

> 💡 The Qualer API key is not stored here — it's securely fetched from `/General/apikey.txt` on SharePoint.

---

## 🛠️ Usage

```bash
python script.py
```

The script:

* Loads `WireSetCerts.xlsx` from `/Shared Documents/Pyro/`.
* Calls the Qualer API to get the most recent cert per asset.
* Uses OCR to extract the wire roll number.
* Saves the updated file back to SharePoint if anything changed.

---

## 🔄 Scheduling

This script is intended to run:

* **Every 10 minutes**.
* **From 5:00 AM to 5:00 PM**.
* **Monday through Saturday**.

You can schedule it using:

* Windows Task Scheduler.
* `cron` (Linux).
* A cloud scheduler (GitHub Actions, Azure Functions, etc.).

> **Clarification**: The script includes an internal loop. It should only be scheduled **once at 5:00 AM**, and will continue running every 10 minutes until 5:00 PM.  It will only attempt to update the Excel file if changes are detected. If no changes are found, the script will exit without making any updates.

---

## 📁 Logging

Logs are written to stdout using Python's `logging` module. Each run logs:

* Current working directory.
* Timestamps.
* Detected changes.
* Errors from SharePoint or Qualer calls.

---

## 📚 Dependencies

```txt
dotenv
git+https://github.com/Johnson-Gage-Inspection-Inc/qualer-sdk-python.git@c948da4#egg=qualer_sdk
msal
openpyxl
pandas
pdf2image
pytesseract
tqdm
```

**Important**
Ensure [Tesseract](https://github.com/tesseract-ocr/tesseract), [Poppler](https://github.com/oschwartz10612/poppler-windows/releases),
and the [Microsoft VC++ Redistributable (x64)](https://aka.ms/vs/17/release/vc_redist.x64.exe) are installed and available in your system path.

## 🧰 Troubleshooting

If you encounter runtime errors related to `pdftoppm`, install the following:

* [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases)
* [Microsoft VC++ Redistributable (x64)](https://aka.ms/vs/17/release/vc_redist.x64.exe)

---

## 📁 File Structure

```
tc-wires/
│
├── script.py                 # Main logic
├── .env                      # Your environment variables (not committed)
├── README.md                 # This file
├── requirements.txt          # Python dependencies
└── task.xml                  # Task Scheduler configuration for Windows
```

---

## 🛡️ Safety

* ✅ Does **not** overwrite unchanged data — compares content hashes first.
* ✅ SharePoint version history preserves previous copies automatically.
* ❌ Excel formulas/macros are **not preserved** — only raw data is written.

---

## 🤝 Contributing

Issues and PRs are welcome. If you improve performance, add PDF pattern support, or switch to cell-level Graph editing, we’d love your input.

---

## 📬 Contact

For support or feedback, contact:

**Jeff Hall**
[jhall@jgiquality.com](mailto:jhall@jgiquality.com)

---

## 📝 License

This project is licensed under the MIT License. See `LICENSE` for details.
