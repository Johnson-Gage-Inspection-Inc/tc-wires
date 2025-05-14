# Wire Set Certificate Processing Script

This script automates the association of **TC wire sets** with the **wire roll** they were made from, by reading each set's calibration certificate (PDF) and extracting the wire roll serial number using OCR. It interacts with the **Qualer SDK**, and **SharePoint** vai the **Microsoft Graph API** to fetch and update data.

---

## 🚀 Features

- 🧠 **Smart Lookups**: Pulls the latest Qualer service record for each wire set and locates the matching certificate document.
- 🧾 **OCR Extraction**: Converts PDFs to images and uses Tesseract to find the wire roll serial number within the certificate.
- 📊 **Excel Updating**: Updates the SharePoint-hosted `WireSetCerts.xlsx` file with extracted data.
- ☁️ **Integrated with SharePoint & Graph API**: Downloads the Excel file and uploads it back only if changes are detected.
- 🔐 **Secure Token Handling**: Pulls the Qualer API token from a protected `apikey.txt` file in SharePoint.
- 🔁 **Scheduled Execution**: Intended to run every 10 minutes from 6:00 AM to 5:00 PM, Monday through Saturday.
- 🪵 **Detailed Logging**: Logs all steps and errors with timestamps using Python's `logging` module.

---

## 🧱 Prerequisites

- Python 3.8+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (installed locally)
- Access to:
  - Microsoft Azure AD App credentials (client ID, secret, tenant ID)
  - Qualer API
  - JGI's SharePoint site

---

## 📦 Installation

```bash
git clone https://github.com/your-org/tc-wires.git
cd tc-wires
pip install -r requirements.txt
````

---

## ⚙️ Configuration

Create a `.env` file in the root directory with:

```env
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=your-app-secret
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
SHAREPOINT_DRIVE_ID=b!xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx  # Documents library drive ID
```

> 💡 The Qualer API key is not stored here — it's securely fetched from `/General/apikey.txt` on SharePoint.

---

## 🛠 Usage

```bash
python script.py
```

The script:

* Loads `WireSetCerts.xlsx` from `/Shared Documents/Pyro/`
* Calls the Qualer API to get the most recent cert per asset
* Uses OCR to extract the wire roll number
* Saves the updated file back to SharePoint if anything changed

---

## 🔄 Scheduling

This script is intended to run:

* **Every 10 minutes**
* **From 6:00 AM to 5:00 PM**
* **Monday through Saturday**

You can schedule it using:

* Windows Task Scheduler
* `cron` (Linux)
* A cloud scheduler (GitHub Actions, Azure Functions, etc.)

---

## 📑 Logging

Logs are written to stdout using Python's `logging` module. Each run logs:

* Current working directory
* Timestamps
* Detected changes
* Errors from SharePoint or Qualer calls

---

## 📚 Dependencies

```txt
pandas
requests
tqdm
pytesseract
pdf2image
python-dotenv
msal
qualer-sdk  # Local or private module
```

Ensure Tesseract and Poppler binaries are installed and available in your system path.

---

## 📁 File Structure

```
tc-wires/
│
├── script.py                 # Main logic
├── .env                      # Your environment variables (not committed)
├── README.md                 # This file
└── requirements.txt          # Python dependencies
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

```

Let me know if you'd like to include:
- A `requirements.txt` scaffold
- Example output logs
- Screenshots of SharePoint integration or OCR results
