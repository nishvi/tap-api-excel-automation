import requests
import pandas as pd
import time

# Your Tap API key
API_KEY = 'sk_live_*********'  # Replace with your key


# File paths
input_file = 'Ooredoocharges.xlsx'              # Your input file with Charge IDs
output_file = 'tap_with_reason_risknew.xlsx'  # Output file with added columns

# Read the Excel file
df = pd.read_excel(input_file)

# Add empty columns
df['Reason'] = ''
df['Risk Rule Name'] = ''

# API headers
headers = {
    'Authorization': f'Bearer {API_KEY}',
    'Content-Type': 'application/json'
}

# Loop through each charge ID
for index, row in df.iterrows():
    charge_id = row['Charge ID']  # Adjust to match your column name
    url = f'https://api.tap.company/v2/charges/{charge_id}'

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()

            # Get reason
            reason = data.get('response', {}).get('message', '')
            df.at[index, 'Reason'] = reason

            # Get first risk rule name (if available)
            rules = data.get('risk', {}).get('rules', [])
            if rules:
                risk_rule_name = rules[0].get('name', '')
                df.at[index, 'Risk Rule Name'] = risk_rule_name
            else:
                df.at[index, 'Risk Rule Name'] = 'No Risk Rule'

        else:
            df.at[index, 'Reason'] = f"Error {response.status_code}"
            df.at[index, 'Risk Rule Name'] = 'API Error'
    except Exception as e:
        df.at[index, 'Reason'] = str(e)
        df.at[index, 'Risk Rule Name'] = 'Exception'

    time.sleep(0.5)  # Respect API limits

# Save final Excel
df.to_excel(output_file, index=False)
print(f"✅ Done! File saved as: {output_file}")