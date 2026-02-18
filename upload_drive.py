import os
import json
import logging
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# Scopes necesarios para subir y gestionar archivos
SCOPES = ['https://www.googleapis.com/auth/drive']
def authenticate_with_service_account():
    """Autentica usando las credenciales de la cuenta de servicio desde una variable de entorno."""
    creds_json = os.environ.get('GDRIVE_SA_KEY')
    if not creds_json:
        raise ValueError("La variable de entorno 'GDRIVE_SA_KEY' no está definida.")
    
    try:
        creds_dict = json.loads(creds_json)
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=SCOPES)
        return creds
    except json.JSONDecodeError:
        raise ValueError("Error al decodificar el JSON de la variable de entorno 'GDRIVE_SA_KEY'.")
def upload_files(folder_id, local_path):
    """Sube archivos desde una ruta local a una carpeta específica de Google Drive."""
    creds = authenticate_with_service_account()
    service = build('drive', 'v3', credentials=creds)
    if not os.path.exists(local_path):
        logging.error(f"La ruta local no existe: {local_path}")
        return
    files_to_upload = []
    if os.path.isfile(local_path):
        files_to_upload.append(local_path)
    elif os.path.isdir(local_path):
        for root, _, files in os.walk(local_path):
            for file in files:
                files_to_upload.append(os.path.join(root, file))
    for file_path in files_to_upload:
        file_name = os.path.basename(file_path)
        logging.info(f"Procesando archivo: {file_name}")
        # Buscar si el archivo ya existe en la carpeta
        query = f"'{folder_id}' in parents and name = '{file_name}' and trashed = false"
        results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        items = results.get('files', [])
        media = MediaFileUpload(file_path, resumable=True)
        if not items:
            # Subir nuevo archivo
            file_metadata = {'name': file_name, 'parents': [folder_id]}
            file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            logging.info(f"Archivo subido: {file_name} (ID: {file.get('id')})")
        else:
            # Actualizar archivo existente
            file_id = items[0]['id']
            file = service.files().update(fileId=file_id, media_body=media, fields='id').execute()
            logging.info(f"Archivo actualizado: {file_name} (ID: {file.get('id')})")
if __name__ == '__main__':
    # ID de la carpeta de destino en Google Drive
    # Puedes obtener esto de la URL de la carpeta: drive.google.com/drive/folders/ESTA_PARTE
    FOLDER_ID = os.environ.get('GDRIVE_FOLDER_ID')
    
    # Ruta local de los archivos a subir
    LOCAL_PATH = os.environ.get('LOCAL_UPLOAD_PATH', './archivos_para_subir')
    if not FOLDER_ID:
        logging.error("La variable de entorno 'GDRIVE_FOLDER_ID' no está definida.")
    else:
        upload_files(FOLDER_ID, LOCAL_PATH)
