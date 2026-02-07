import requests
import os
from datetime import datetime, timedelta
import calendar
import locale
try:
    import pandas as pd
except ImportError:
    pd = None
    print("Advertencia: pandas no está instalado. No se podrá procesar archivos TIE.")

# ... (resto de imports y constantes igual)


try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
    except locale.Error:
        print("Advertencia: No se pudo establecer locale en español. Los nombres de carpeta podrían ser incorrectos.")

# Configuración de Esquemas y Archivos
ESQUEMAS = {
    "Mensual": {
        "carpeta_url": "Energia y Mercado/Garantias Mensuales",
        "archivos": [
            "GARANTIA SEMANAL MENSUAL",
            "GARANTIA TXR",
            "GARANTIA MENSUAL",
            "GARANTIA CUOTA RES 101 029 2022"
        ]
    },
    "Semanal": {
        "carpeta_url": "Energia y Mercado/Garantias Semanales",
        "archivos": [
            "GARANTIA SEMANAL",
            "GARANTIA TXR"
        ]
    },
    "TIE": {
        "carpeta_url": "Agentes/Garantias Financieras TIE",
        "archivos": [
            "WEB_GARANTIES",
            "WEB_GARANTIAS"
        ],
        "formato_fecha": "NUMERICO", # DD-MM-YYYY
        "separador": "-", # WEB_GARANTIES-10...
        "formato_carpeta_mes": "SIN_ESPACIO" # 02.Febrero (vs 02. Febrero)
    },
    "Cuentas": {
        "carpeta_url": "Agentes/SaldosDiariosCuentasCustodia",
        "archivos": [
            "Saldo cuenta custodia"
        ],
        "formato_fecha": "ISO", # YYYY-MM-DD
        "separador": " ",
        "incluir_path_fecha": False # No usa subcarpetas de año/mes
    }
}

def get_xm_url(filename_base, date_obj, esquema_nombre="Mensual", version_suffix="", extension=".xlsx"):
    """
    Genera la URL de descarga basándose en el esquema y fecha.
    version_suffix: Sufijo opcional (ej: "_V2") que se agrega antes de la extensión.
    extension: Extensión del archivo (ej: ".xlsx", ".XLSX").
    """
    
    # Obtener configuración del esquema
    if esquema_nombre in ESQUEMAS:
        config = ESQUEMAS[esquema_nombre]
        carp_garantias = config["carpeta_url"]
        formato_fecha = config.get("formato_fecha", "TEXTO") # TEXTO (23ENE), NUMERICO (23-01), ISO (2026-02-06)
        separador = config.get("separador", " ")
        formato_carpeta_mes = config.get("formato_carpeta_mes", "CON_ESPACIO")
        incluir_path_fecha = config.get("incluir_path_fecha", True)
    else:
        # Fallback (asumimos lógica Mensual/Semanal antigua)
        carp_garantias = "Energia y Mercado/Garantias Mensuales"
        if "SEMANAL" in filename_base and "MENSUAL" not in filename_base:
            carp_garantias = "Energia y Mercado/Garantias Semanales"
        formato_fecha = "TEXTO"
        separador = " "
        formato_carpeta_mes = "CON_ESPACIO"
        incluir_path_fecha = True

    # Mapeo manual de meses...
    meses_nombres = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
        7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
    }
    
    meses_abrev = {
        1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
        7: "JUL", 8: "AGO", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC"
    }

    # 1. Carpeta Mes (Solo si aplica)
    month_num = date_obj.strftime("%m")
    year_str = date_obj.strftime("%Y")
    day_str = date_obj.strftime("%d")
    
    folder_path_part = ""
    if incluir_path_fecha:
        month_name = meses_nombres[date_obj.month]
        if formato_carpeta_mes == "SIN_ESPACIO":
            folder_month = f"{month_num}.{month_name}"
        else:
            folder_month = f"{month_num}. {month_name}"
        folder_path_part = f"/{year_str}/{folder_month}"
    
    # 2. Nombre Archivo y Fecha
    
    if formato_fecha == "NUMERICO":
        # Formato: 10-02-2026
        file_date_str = f"{day_str}-{month_num}-{year_str}"
    elif formato_fecha == "ISO":
        # Formato: 2026-02-06
        file_date_str = f"{year_str}-{month_num}-{day_str}"
    else:
        # Formato: 23ENE-2026
        month_abrev = meses_abrev[date_obj.month]
        file_date_str = f"{day_str}{month_abrev}-{year_str}"
    
    # Construir nombre completo
    # Ej: NOMBRE + " " + FECHA + _V2 + .xlsx
    full_filename = f"{filename_base}{separador}{file_date_str}{version_suffix}{extension}"
    
    
    # 3. URL
    base_url = "https://app-portalxmcore01.azurewebsites.net/administracion-archivos/ficheros/descarga-archivo"
    # Nota: Ya no agregamos "Energia y Mercado/" duro, viene en carp_garantias
    # folder_path_part ya incluye timestamps si son necesarios, o es vacio si flat
    blob_path = f"{carp_garantias}{folder_path_part}/{full_filename}"
    
    params = {
        'ruta': blob_path,
        'nombreBlobContainer': 'storageportalxm'
    }
    
    req = requests.Request('GET', base_url, params=params)
    prepped = req.prepare()
    return prepped.url, full_filename

def download_file(url, filename, save_dir="Descargas_XM"):
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
        
    save_path = os.path.join(save_dir, filename)
    
    print(f"Descargando: {filename}...")
    print(f"URL: {url}")
    
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status() # Lanza error si es 404, 403, etc.
        
        with open(save_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        print(f"¡Éxito! Guardado en: {save_path}")
        return True
    except requests.exceptions.HTTPError as err:
        print(f"Error HTTP: {err}")
        return False
    except Exception as e:
        print(f"Error general: {e}")
        return False

def clean_tie_file(filepath):
    """
    Elimina la primera columna del archivo Excel dado (TIE).
    Sobreescribe el archivo original.
    Retorna: (nuevo_path, error_msg)
    """
    if pd is None:
        return None, "Librería pandas no instalada."
        
    try:
        print(f"Procesando TIE: Eliminando primera columna de {filepath}...")
        
        # Detectar motor según extensión
        engine = 'openpyxl' if filepath.lower().endswith('.xlsx') else 'xlrd'
        
        # Leer archivo
        try:
            df = pd.read_excel(filepath, engine=engine)
        except Exception as e_read:
            # Fallback: intentar el otro motor si falló (a veces extensión miente)
            alt_engine = 'xlrd' if engine == 'openpyxl' else 'openpyxl'
            try:
                print(f"Reintentando con motor alternativo {alt_engine}...")
                df = pd.read_excel(filepath, engine=alt_engine)
            except Exception as e_retry:
                return None, f"Error leyendo Excel ({e_read}) | Retry ({e_retry})"

        # Eliminar primera columna (index 0)
        if not df.empty:
            df.drop(df.columns[0], axis=1, inplace=True)
        
        # Guardar (sin index)
        output_path = filepath
        # Siempre intentamos guardar como xlsx moderno para evitar lios
        if filepath.lower().endswith('.xls'):
            output_path = filepath + "x" 
            
        df.to_excel(output_path, index=False)
        
        if output_path != filepath:
            try:
                os.remove(filepath)
            except:
                pass # Si no se puede borrar el viejo, no es crítico
            print(f"Archivo actualizado a formato moderno: {output_path}")
            return output_path, None
            
        print("¡Archivo TIE procesado correctamente!")
        return output_path, None
        
    except Exception as e:
        print(f"Error procesando TIE: {e}")
        return None, str(e)

if __name__ == "__main__":
    import sys
    # ... (Actualizar main para soportar la nueva lógica si se ejecuta directo es opcional, 
    # pero nos enfocaremos en que sirva de librería para la GUI) ...
    pass
