import os

import requests

tenant_id = os.environ["SETTINGS__GRAPH_API__TENANT_ID"]
application_id = os.environ["SETTINGS__GRAPH_API__CLIENT_ID"]
client_secret = os.environ["SETTINGS__GRAPH_API__CLIENT_SECRET"]

# get access token (should probably use a proper OAuth2 client for this)
response = requests.post(
    f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
    data={
        "client_id": application_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    },
)
response.raise_for_status()

access_token = response.json()
token = access_token["access_token"]

with requests.Session() as session:
    session.headers.update({"Authorization": f"Bearer {token}"})

    # For now use a fake submission
    test_submission = {
        "reference": "AAAAAA",
        "submitted_at": "2026-01-26T10:48:00Z",
        "What’s your name?": "Form Filler",
        "When’s your date of birth?": "1990-01-01",
        "What’s your address?": "1 Fake Street, Notatown, AA1 2AA",
    }

    breakpoint()
