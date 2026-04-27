#!/usr/bin/env python3
"""Exchange OAuth code for token and upload files to Google Drive."""
import json
import os
import sys

CREDS_FILE = '/root/.claude/gdrive-credentials.json'
TOKEN_FILE = '/root/.claude/gdrive-token.json'

CODE = '4/1AeoWuM_mOIzqFZtHeq0xyeRRQycow-6kSDXagWzOyTulxinhU7ulbxsB7pk'

def exchange_token():
    from google_auth_oauthlib.flow import Flow
    SCOPES = ['https://www.googleapis.com/auth/drive']
    with open(CREDS_FILE) as f:
        raw = json.load(f)
    if 'installed' in raw:
        client_config = raw
    else:
        client_config = {'installed': {**raw, 'redirect_uris': ['urn:ietf:wg:oauth:2.0:oob']}}
    flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri='urn:ietf:wg:oauth:2.0:oob')
    flow.fetch_token(code=CODE)
    creds = flow.credentials
    with open(TOKEN_FILE, 'w') as f:
        f.write(creds.to_json())
    print(f"✓ Token guardado en {TOKEN_FILE}")
    return creds

def main():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload

    OUTPUT_DIR = '/home/user/sistema-dies/outputs'
    FOLDER_NAME = 'Sistema Comercial Operativo — León López'
    FILES = [
        'diagnostico_cuestionario.docx',
        'diagnostico_calculadora_oportunidad.xlsx',
        'diagnostico_informe_cliente.docx',
        'implementacion_playbook_operativo.docx',
        'implementacion_acuerdo_cliente.docx',
        'ejecucion_plan_trabajo_8semanas.xlsx',
        'ejecucion_protocolo_reunion_semanal.docx',
        'ejecucion_protocolo_seguimiento_leads.docx',
        'socializacion_manual_dueno.docx',
        'socializacion_kit_lanzamiento_equipo.docx',
        'socializacion_reporte_mensual_template.docx',
    ]
    MIME_TYPES = {
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }

    print("Obteniendo token de autorización...")
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, ['https://www.googleapis.com/auth/drive'])
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                creds = exchange_token()
    else:
        creds = exchange_token()

    service = build('drive', 'v3', credentials=creds)

    # Create or find folder
    print(f"\nCreando carpeta '{FOLDER_NAME}'...")
    query = f"name='{FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, fields='files(id, name)').execute()
    items = results.get('files', [])
    if items:
        folder_id = items[0]['id']
        print(f"  Carpeta existente: {folder_id}")
    else:
        metadata = {'name': FOLDER_NAME, 'mimeType': 'application/vnd.google-apps.folder'}
        folder = service.files().create(body=metadata, fields='id').execute()
        folder_id = folder['id']
        print(f"  Carpeta creada: {folder_id}")

    print(f"\nSubiendo {len(FILES)} archivos...\n")
    uploaded = []
    errors = []
    for filename in FILES:
        filepath = os.path.join(OUTPUT_DIR, filename)
        if not os.path.exists(filepath):
            print(f"  ✗ No encontrado: {filename}")
            errors.append(filename)
            continue
        try:
            ext = os.path.splitext(filename)[1].lower()
            mime_type = MIME_TYPES.get(ext, 'application/octet-stream')
            metadata = {'name': filename, 'parents': [folder_id]}
            media = MediaFileUpload(filepath, mimetype=mime_type, resumable=True)
            result = service.files().create(body=metadata, media_body=media, fields='id,name,webViewLink').execute()
            link = result.get('webViewLink', f"https://drive.google.com/file/d/{result['id']}")
            print(f"  ✓ {filename}")
            print(f"    {link}")
            uploaded.append({'name': filename, 'link': link})
        except Exception as e:
            print(f"  ✗ Error con {filename}: {e}")
            errors.append(filename)

    folder_link = f"https://drive.google.com/drive/folders/{folder_id}"
    print(f"\n{'='*60}")
    print(f"Subidos: {len(uploaded)}/{len(FILES)}")
    if errors:
        print(f"Errores: {errors}")
    print(f"\nCarpeta en Google Drive:")
    print(f"  {folder_link}")

if __name__ == '__main__':
    main()
