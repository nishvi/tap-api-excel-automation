# Tap Payments Charge API Automation

This script reads charge IDs from an Excel file, calls Tap Payments API for each charge, and outputs a new Excel file with:
- Response Reason
- Risk Rule Name

## Technologies
- Python
- Pandas
- Requests
- OpenPyXL

## How to Use
1. Replace the API key in the script
2. Place your input file as `charges.xlsx`
3. Run the script: `python tap_reasons.py`

## Output
Generates `tap_with_reason_risk.xlsx` with all details.

*Note: Do not upload real API keys or sensitive data.*

---
