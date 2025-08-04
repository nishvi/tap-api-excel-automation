import requests
import pandas as pd
import time

# Your Tap API key
API_KEY = 'sk_live_*********'  # Replace with your key


# File paths
input_file = 'Ooredoocharges.xlsx'
output_file = 'RISK2.xlsx'

# Read the Excel file
df = pd.read_excel(input_file)

# Add empty columns
df['Reason'] = ''
df['Risk Rule Name'] = ''
df['Merchant ID'] = ''
df['Payment Method'] = ''
df['Issuer Country'] = ''
df['Issuer Region Code'] = ''
df['Issuer Region Name'] = ''

# API headers
headers = {
    'Authorization': f'Bearer {API_KEY}',
    'Content-Type': 'application/json'
}

# Loop through each charge ID
for index, row in df.iterrows():
    charge_id = row['Charge ID']
    url = f'https://api.tap.company/v2/charges/{charge_id}'

    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()

            # Reason
            reason = data.get('response', {}).get('message', '')
            df.at[index, 'Reason'] = reason

            # Risk Rule Names - all rule names joined by comma
            risk_data = data.get('risk', {})
            rules = risk_data.get('rules', [])
            if isinstance(rules, list) and rules:
                rule_names = ', '.join([rule.get('name', '') for rule in rules if 'name' in rule])
                df.at[index, 'Risk Rule Name'] = rule_names
            else:
                df.at[index, 'Risk Rule Name'] = 'No Rule Found'

            # Merchant ID
            merchant_id = data.get('merchant', {}).get('id', '')
            df.at[index, 'Merchant ID'] = merchant_id

            # Payment Method
            payment_method = data.get('source', {}).get('payment_method', '')
            df.at[index, 'Payment Method'] = payment_method

            # Issuer details
            issuer = data.get('card', {}).get('issuer', {})
            issuer_country = issuer.get('country', '')
            issuer_region = issuer.get('issuer_region', {})

            df.at[index, 'Issuer Country'] = issuer_country
            df.at[index, 'Issuer Region Code'] = issuer_region.get('code', '')
            df.at[index, 'Issuer Region Name'] = issuer_region.get('name', '')

        else:
            # API returned error status
            for col in ['Reason', 'Risk Rule Name', 'Merchant ID', 'Payment Method',
                        'Issuer Country', 'Issuer Region Code', 'Issuer Region Name']:
                df.at[index, col] = f"Error {response.status_code}"

    except Exception as e:
        # Exception occurred
        for col in ['Reason', 'Risk Rule Name', 'Merchant ID', 'Payment Method',
                    'Issuer Country', 'Issuer Region Code', 'Issuer Region Name']:
            df.at[index, col] = str(e)

    time.sleep(0.5)

# Save to Excel
df.to_excel(output_file, index=False)
print(f"✅ Done! File saved as: {output_file}")
