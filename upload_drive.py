import os
import json
import logging
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# Scopes necesarios
SCOPES = ['https://www.googleapis.com/auth/drive']
def authenticate_with_token_json():
    """Autentica usando el token.json guardado en la variable de entorno."""
    token_content = os.environ.get('GDRIVE_TOKEN_JSON')
    if not token_content:
        raise ValueError("La variable de entorno 'GDRIVE_TOKEN_JSON' no está definida.")
    
    try:
        # Convertimos la string JSON a un diccionario
        token_dict = json.loads(token_content)
        # Reconstruimos las credenciales
        creds = Credentials.from_authorized_user_info(token_dict, SCOPES)
        return creds
    except json.JSONDecodeError:
        raise ValueError("Error al decodificar el JSON de la variable de entorno 'GDRIVE_TOKEN_JSON'.")
def get_or_create_folder(service, parent_id, folder_name):
    """Busca una carpeta dentro de parent_id por nombre, si no existe la crea."""
    query = f"'{parent_id}' in parents and name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    items = results.get('files', [])
    
    if items:
        folder_id = items[0]['id']
        logging.info(f"Carpeta existente encontrada: {folder_name} (ID: {folder_id})")
        return folder_id
    else:
        logging.info(f"Creando carpeta: {folder_name}")
        file_metadata = {
            'name': folder_name,
            'parents': [parent_id],
            'mimeType': 'application/vnd.google-apps.folder'
        }
        folder = service.files().create(body=file_metadata, fields='id').execute()
        logging.info(f"Carpeta creada: {folder_name} (ID: {folder.get('id')})")
        return folder.get('id')
def upload_files_recursive(base_folder_id, local_root):
    """Sube archivos recursivamente manteniendo la estructura de carpetas."""
    try:
        creds = authenticate_with_token_json()
        service = build('drive', 'v3', credentials=creds)
    except Exception as e:
        logging.error(f"Error de autenticación: {e}")
        return
    if not os.path.exists(local_root):
        logging.error(f"La ruta local no existe: {local_root}")
        return
    # Diccionario para cachear los IDs de las carpetas creadas/encontradas
    # Clave: ruta relativa (ej: "Semanal"), Valor: drive_folder_id
    folder_cache = {}
    for root, dirs, files in os.walk(local_root):
        # Calcular ruta relativa desde la carpeta raíz local
        if root == local_root:
            rel_path = "."
        else:
            rel_path = os.path.relpath(root, local_root)
        
        # Determinar el ID de la carpeta padre en Drive para este nivel
        if rel_path == '.':
            current_drive_folder_id = base_folder_id
        else:
            # Descomponer la ruta relativa (ej: "2024/Enero")
            parts = rel_path.split(os.sep)
            parent_id = base_folder_id
            current_path_built = ""
            
            # Recorrer/Crear estructura de carpetas nivel por nivel
            for part in parts:
                if current_path_built:
                    current_path_built = os.path.join(current_path_built, part)
                else:
                    current_path_built = part
                
                # Chequear cache para no llamar a la API innecesariamente
                if current_path_built in folder_cache:
                    parent_id = folder_cache[current_path_built]
                else:
                    # Buscar o crear en Drive
                    parent_id = get_or_create_folder(service, parent_id, part)
                    folder_cache[current_path_built] = parent_id
            
            current_drive_folder_id = parent_id
        # Subir archivos en la carpeta actual
        for file_name in files:
            local_file_path = os.path.join(root, file_name)
            display_path = os.path.join(rel_path, file_name) if rel_path != "." else file_name
            logging.info(f"Procesando archivo: {display_path}")
            
            try:
                # Buscar si el archivo ya existe en la carpeta destino
                query = f"'{current_drive_folder_id}' in parents and name = '{file_name}' and trashed = false"
                results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
                items = results.get('files', [])
                media = MediaFileUpload(local_file_path, resumable=True)
                if not items:
                    file_metadata = {'name': file_name, 'parents': [current_drive_folder_id]}
                    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                    logging.info(f"Subido: {file_name}")
                else:
                    file_id = items[0]['id']
                    file = service.files().update(fileId=file_id, media_body=media, fields='id').execute()
                    logging.info(f"Actualizado: {file_name}")
            
            except Exception as e:
                logging.error(f"Error al subir/actualizar {file_name}: {e}")
if __name__ == '__main__':
    # ID de la carpeta raíz de destino en Google Drive
    FOLDER_ID = os.environ.get('GDRIVE_FOLDER_ID')
    
    # Ruta local de los archivos a subir (ej: ./Garantías)
    LOCAL_PATH = os.environ.get('LOCAL_UPLOAD_PATH', './Garantías')
    if not FOLDER_ID:
        logging.error("La variable de entorno 'GDRIVE_FOLDER_ID' no está definida.")
    else:
        FOLDER_ID = FOLDER_ID.strip()
        upload_files_recursive(FOLDER_ID, LOCAL_PATH)
