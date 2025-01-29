import streamlit as st
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/contacts']

# Access secrets
google_secrets = st.secrets["google"]

client_id = google_secrets["client_id"]
client_secret = google_secrets["client_secret"]
redirect_uri = google_secrets["redirect_uri"]
project_id = google_secrets["project_id"]
auth_uri = google_secrets["auth_uri"]
token_uri = google_secrets["token_uri"]
auth_provider_x509_cert_url = google_secrets["auth_provider_x509_cert_url"]

def authenticate_google():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Use secrets from st.secrets
            client_config = {
                "installed": {
                    "client_id": st.secrets["google"]["client_id"],
                    "client_secret": st.secrets["google"]["client_secret"],
                    "redirect_uris": [st.secrets["google"]["redirect_uri"]],
                    "auth_uri": st.secrets["google"]["auth_uri"],
                    "token_uri": st.secrets["google"]["token_uri"],
                }
            }
            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def add_contact(service, name, phone):
    contact = {
        'names': [{'givenName': name}],
        'phoneNumbers': [{'value': str(phone)}]
    }
    service.people().createContact(body=contact).execute()

def main():
    st.title("Excel to Google Contacts Converter")
    
    # Upload Excel file
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        # Load the Excel file
        df = pd.read_excel(uploaded_file)
        
        # Display the data
        st.write("Preview of the uploaded data:")
        st.dataframe(df)
        
        # Extract 'Name' and 'Phone' columns
        if 'Name' in df.columns and 'Phone' in df.columns:
            contacts = df[['Name', 'Phone']].values.tolist()
            
            # Authenticate and build the service
            if st.button("Add Contacts to Google"):
                with st.spinner("Authenticating and adding contacts..."):
                    creds = authenticate_google()
                    service = build('people', 'v1', credentials=creds)
                    
                    # Loop through contacts and add them
                    for contact in contacts:
                        name, phone = contact
                        add_contact(service, name, phone)
                    st.success("Contacts added successfully!")
        else:
            st.error("The Excel file must contain 'Name' and 'Phone' columns.")

if __name__ == '__main__':
    main() 