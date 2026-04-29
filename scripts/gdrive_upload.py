#!/usr/bin/env python3
"""Upload Sistema Comercial Operativo files to Google Drive."""
import os
import json
import sys
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ['https://www.googleapis.com/auth/drive']
CREDS_FILE = '/root/.claude/gdrive-credentials.json'
TOKEN_FILE = '/root/.claude/gdrive-token.json'

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


def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Load the client secrets
            with open(CREDS_FILE) as f:
                raw = json.load(f)
            # Normalize: supports both direct format and {"installed": {...}} format
            if 'installed' in raw:
                client_config = raw
            else:
                client_config = {'installed': raw}
            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            # Use run_local_server with a fixed port
            creds = flow.run_local_server(port=8080, open_browser=False)
        with open(TOKEN_FILE, 'w') as f:
            f.write(creds.to_json())
    return creds


def create_folder(service, name):
    query = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, fields='files(id, name)').execute()
    items = results.get('files', [])
    if items:
        print(f"  Carpeta existente encontrada: {items[0]['id']}")
        return items[0]['id']
    metadata = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    folder = service.files().create(body=metadata, fields='id').execute()
    print(f"  Carpeta creada: {folder['id']}")
    return folder['id']


def upload_file(service, filepath, folder_id):
    filename = os.path.basename(filepath)
    ext = os.path.splitext(filename)[1].lower()
    mime_type = MIME_TYPES.get(ext, 'application/octet-stream')
    metadata = {'name': filename, 'parents': [folder_id]}
    media = MediaFileUpload(filepath, mimetype=mime_type, resumable=True)
    file = service.files().create(
        body=metadata, media_body=media, fields='id, name, webViewLink'
    ).execute()
    return file


def main():
    print("Autenticando con Google Drive...")
    creds = get_credentials()
    service = build('drive', 'v3', credentials=creds)

    print(f"\nCreando carpeta '{FOLDER_NAME}'...")
    folder_id = create_folder(service, FOLDER_NAME)

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
            print(f"  ↑ {filename}...")
            result = upload_file(service, filepath, folder_id)
            print(f"    ✓ Subido: {result.get('webViewLink', result['id'])}")
            uploaded.append(result)
        except Exception as e:
            print(f"    ✗ Error: {e}")
            errors.append(filename)

    print(f"\n{'='*50}")
    print(f"Subidos: {len(uploaded)} / {len(FILES)}")
    if errors:
        print(f"Errores: {errors}")
    folder_link = f"https://drive.google.com/drive/folders/{folder_id}"
    print(f"\nCarpeta en Drive: {folder_link}")
    return 0 if not errors else 1


if __name__ == '__main__':
    sys.exit(main())
