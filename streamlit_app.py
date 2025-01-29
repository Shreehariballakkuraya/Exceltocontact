import streamlit as st
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow, Flow
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
                "web": {
                    "client_id": st.secrets["google"]["client_id"],
                    "client_secret": st.secrets["google"]["client_secret"],
                    "auth_uri": st.secrets["google"]["auth_uri"],
                    "token_uri": st.secrets["google"]["token_uri"],
                    "redirect_uris": [
                        "https://exceltocontacts.streamlit.app",
                        "https://exceltocontacts.streamlit.app/",
                        "http://localhost:8501",
                        "http://localhost:8501/"
                    ],
                    "javascript_origins": [
                        "https://exceltocontacts.streamlit.app",
                        "http://localhost:8501"
                    ]
                }
            }
            
            flow = Flow.from_client_config(
                client_config,
                scopes=SCOPES,
                redirect_uri="https://exceltocontacts.streamlit.app"
            )
            
            # Generate the authorization URL
            auth_url, _ = flow.authorization_url(
                access_type='offline',
                include_granted_scopes='true',
                prompt='consent'
            )
            
            # Display instructions
            st.markdown("""
            ### Authorization Steps:
            1. Click the link below to authorize the application
            2. After authorizing, you'll be redirected to a page that might show an error - this is expected
            3. Copy the entire URL from your browser's address bar after being redirected
            4. Paste the URL in the text box below
            """)
            
            st.markdown(f"[Click here to authorize the application]({auth_url})")
            
            auth_response = st.text_input("Paste the full redirect URL here:")
            
            if auth_response:
                try:
                    # Extract the authorization code from the full redirect URL
                    from urllib.parse import urlparse, parse_qs
                    parsed = urlparse(auth_response)
                    auth_code = parse_qs(parsed.query)['code'][0]
                    
                    flow.fetch_token(code=auth_code)
                    creds = flow.credentials
                    with open('token.json', 'w') as token:
                        token.write(creds.to_json())
                except Exception as e:
                    st.error(f"Error during authentication: {str(e)}")
                    return None
            else:
                st.info("Please paste the redirect URL to complete authentication.")
                return None
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
                    if creds:
                        service = build('people', 'v1', credentials=creds)
                        
                        # Loop through contacts and add them
                        for contact in contacts:
                            name, phone = contact
                            add_contact(service, name, phone)
                        st.success("Contacts added successfully!")
                    else:
                        st.error("Authentication failed. Please try again.")
        else:
            st.error("The Excel file must contain 'Name' and 'Phone' columns.")

if __name__ == '__main__':
    main() 