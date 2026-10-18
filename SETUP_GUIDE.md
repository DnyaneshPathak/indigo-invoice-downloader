# IndiGo GST Invoice Downloader — Setup & User Guide

This guide explains how to set up and run this automation script on any Windows system (or transfer to a new machine).

---

## 1. Prerequisites on the New System

1. **Python 3.10 or higher**:
   - Download from [python.org](https://www.python.org/downloads/).
   - ⚠️ **Important during installation**: Check the box **"Add Python to PATH"**.

2. **Google Chrome**:
   - Install the official [Google Chrome](https://www.google.com/chrome/) browser.
   - Ensure Chrome is up to date (Chrome Menu `⋮` → *Help* → *About Google Chrome*).

---

## 2. Step-by-Step Installation

### Step A: Open PowerShell or Command Prompt
Open the folder containing `main.py` in File Explorer, click the address bar, type `cmd` or `powershell`, and press **Enter**.

### Step B: (Recommended) Create a Virtual Environment
```powershell
python -m venv venv
.\venv\Scripts\activate
```

### Step C: Install Required Libraries
Run:
```powershell
pip install -r requirements.txt
```

> **Manual install alternative:**  
> If you prefer installing packages manually without `requirements.txt`:  
> `pip install undetected-chromedriver selenium openpyxl pandas`

---

## 3. Preparing Input Data

Place your input file directly inside this folder. The script auto-detects any of the following names:
- `pnr.xlsx` / `pnr.xls`
- `pnr.csv`
- `input.xlsx` / `input.csv`

### Supported File Format:
Your file should have two main columns (headers are auto-detected):
| PNR / Ticket No | Email ID / GST Email |
| --------------- | -------------------- |
| ABCDEF          | company@domain.com   |
| XYZ123          | gst@domain.com       |

*Note: If no headers are present, Column 1 is read as PNR, and Column 2 is read as Email.*

---

## 4. How to Run the Automation

Run the script using:
```powershell
python main.py
```

### What Happens When It Runs:
1. **Pre-warm Check**: Tests Chrome and patches the driver once before launching.
2. **Resume Check**: If `output.csv` exists, it automatically skips PNRs that have already been downloaded successfully.
3. **Automated Browser**: Opens Chrome, inputs the PNR and GST email, expands "Get Your GST" to extract all invoices under that PNR, and downloads each invoice PDF.

---

## 5. Built-in Safety & Protection Features

To prevent IndiGo (Akamai) from temporarily blocking your IP, the script has built-in protections:

- **Batch Processing**: Processes **90 PNRs per batch**, then pauses for **10 minutes** to allow the site's rate-limit window to reset.
- **Circuit Breaker**: If **5 consecutive PNRs fail** in a row, the script enters a **30-minute sleep** cooldown and restarts the browser.
- **Human-like Delays**: Waits randomized delays (5–12 seconds) between successful PNRs and types keystrokes with slight random variance.
- **Auto-Refresh Mechanism**: If IndiGo shows a transient *"Unable to fetch"* response, it refreshes the page and retries once automatically.
- **Periodic Browser Recycle**: Restarts Chrome every 20 PNRs to clear memory cache and avoid browser slowdowns.

---

## 6. Output Files & Directories

- **`invoices/`**: Directory where all downloaded invoice PDFs are saved (`<InvoiceNumber>.pdf`).
- **`output.csv`**: Records successful PNRs and their downloaded invoice numbers:
  ```csv
  PNR,INVOICES
  ABCDEF,INV100234,INV100235
  ```
- **`pnr_errors.csv`**: Logs any PNR that failed after all retries with reason and timestamp:
  ```csv
  PNR,ERROR,TIME
  XYZ999,Site error: invalid details,22:15:30
  ```

---

## 7. Troubleshooting & Common Questions

### 1. `Driver pre-warm failed` or Chrome version mismatch:
If Chrome updated recently or driver binary got corrupted:
1. Press `Win + R`, type `%APPDATA%`, and press Enter.
2. Find the folder named `undetected_chromedriver` and delete it.
3. Re-run `python main.py` (it will auto-download the correct matching driver).

### 2. Can I stop and resume later?
**Yes.** Press `Ctrl + C` in the terminal to stop anytime. When you run `python main.py` again, `RESUME = True` will check `output.csv` and only process the remaining pending PNRs.

### 3. Rate Limit / "Unable to fetch your GST invoice right now":
- IndiGo's firewall (Akamai) limits requests per IP.
- The script handles this automatically via the **10-minute batch rest** and **30-minute circuit breaker**.
- For optimal results, running outside peak daytime hours (e.g., evening or overnight) offers higher success rates.
