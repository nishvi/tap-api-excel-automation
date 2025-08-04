# Tap Payments Charge API Automation

This Python script automates the process of retrieving charge details from the Tap Payments API. It reads charge IDs from an Excel file, calls the API for each, and outputs a new Excel file with rich risk analysis and issuer metadata.

## 🔍 Features
- Reads charge IDs from an input Excel file
- Fetches and records:
  - Response reason
  - Merchant ID
  - Payment method
  - Issuer country
  - Issuer region (code & name)
  - Up to 5 risk rule details:
    - Rule Name
    - Risk Level
    - Decision
    - Status

## 🛠 Technologies Used
- Python 3
- `pandas`
- `requests`
- `openpyxl` (for Excel output)

## 📦 How to Use

1. **Replace the API key** in the script with your Tap secret key.
2. **Place your input Excel file** in the same directory with column `Charge ID` (e.g., `Ooredoocharges.xlsx`).
3. **Run the script**:
   ```bash
   python tap_reasons.py
