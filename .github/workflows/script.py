import requests
import pandas as pd
import datetime
import os

# Step 1: Call API
api_url = "https://formatter.org/"  # adjust endpoint if needed
response = requests.get(api_url)

if response.status_code != 200:
    raise Exception("API failed: " + str(response.status_code))

data = response.json()

# Step 2: Convert to CSV
df = pd.json_normalize(data)

timestamp = datetime.datetime.utcnow().strftime("%Y-%m-%d_%H-%M-%S")
file_name = f"timesheet_{timestamp}.csv"

df.to_csv(file_name, index=False)

# Step 3: Get Access Token (Azure AD)
token_url = f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}/oauth2/v2.0/token"

token_data = {
    "grant_type": "client_credentials",
    "client_id": os.environ["CLIENT_ID"],
    "client_secret": os.environ["CLIENT_SECRET"],
    "scope": "https://graph.microsoft.com/.default"
}

token_response = requests.post(token_url, data=token_data)
access_token = token_response.json()["access_token"]

# Step 4: Upload to SharePoint
headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "text/csv"
}

upload_url = f"https://graph.microsoft.com/v1.0/sites/{os.environ['SITE_ID']}/drives/{os.environ['DRIVE_ID']}/root:/timesheets/{file_name}:/content"

with open(file_name, "rb") as f:
    upload_response = requests.put(upload_url, headers=headers, data=f)

if upload_response.status_code not in [200, 201]:
    raise Exception(upload_response.text)

print("File uploaded successfully")
