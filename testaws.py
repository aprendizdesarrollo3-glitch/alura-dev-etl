# test_secret.py
from gmail.client import get_gmail_service

service = get_gmail_service()

results = service.users().messages().list(
    userId="me",
    maxResults=5
).execute()

print(results.get("messages", []))