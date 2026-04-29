#!/usr/bin/env python3
"""Step 1: Generate Google OAuth URL for Drive access."""
import json
import sys
from google_auth_oauthlib.flow import Flow

SCOPES = ['https://www.googleapis.com/auth/drive']
CREDS_FILE = '/root/.claude/gdrive-credentials.json'
TOKEN_FILE = '/root/.claude/gdrive-token.json'
REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob'  # out-of-band (copy/paste code)

def main():
    with open(CREDS_FILE) as f:
        raw = json.load(f)
    if 'installed' in raw:
        client_config = raw
    else:
        client_config = {'installed': {**raw, 'redirect_uris': ['urn:ietf:wg:oauth:2.0:oob']}}

    flow = Flow.from_client_config(
        client_config,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )
    auth_url, _ = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='consent'
    )
    print("\n" + "="*60)
    print("PASO 1: Visita esta URL en tu navegador:")
    print("="*60)
    print(auth_url)
    print("="*60)
    print("\nDespués de autorizar, Google te dará un código.")
    print("Copia ese código y pégalo aquí:")
    code = input("Código de autorización: ").strip()

    flow.fetch_token(code=code)
    creds = flow.credentials
    with open(TOKEN_FILE, 'w') as f:
        f.write(creds.to_json())
    print(f"\n✓ Token guardado en {TOKEN_FILE}")
    print("Ahora puedes ejecutar: python3 gdrive_upload.py")

if __name__ == '__main__':
    main()
