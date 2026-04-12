import requests
import os
import time
from datetime import datetime, timedelta
import calendar
import locale
import ssl
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import pandas as pd
except ImportError:
    pd = None
    print("Advertencia: pandas no está instalado. No se podrá procesar archivos TIE.")

# Configuración de Sesión Global con reintentos automáticos ante errores SSL/conexión
session = requests.Session()
retry_strategy = Retry(
    total=3,
    backoff_factor=2,          # espera 2s, 4s, 8s entre reintentos
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["GET"],
)
adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=5, pool_maxsize=5)
session.mount('https://', adapter)
session.mount('http://', adapter)

# Headers para parecer un navegador y evitar bloqueos por User-Agent
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'application/octet-stream, */*',
    'Accept-Language': 'es-CO,es;q=0.9',
})


import concurrent.futures

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

    for attempt in range(3):
        try:
            response = session.get(url, stream=True, timeout=45)
            response.raise_for_status()

            with open(save_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"¡Éxito! Guardado en: {save_path}")
            return True
        except requests.exceptions.HTTPError:
            # 404/403 → el archivo no existe, no reintentar
            return False
        except (requests.exceptions.SSLError, requests.exceptions.ConnectionError) as e:
            if attempt < 2:
                time.sleep(3 * (attempt + 1))  # espera 3s, 6s antes de reintentar
                continue
            print(f"Error SSL/conexión en {filename}: {e}")
            return False
        except Exception as e:
            print(f"Error general en {filename}: {e}")
            return False
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


# --- LOGICA DE VERIFICACIÓN Y REPORTES ---

def read_maestro_file(filepath):
    """
    Lee el archivo Maestro y retorna una lista de diccionarios con la info de los agentes.
    Columnas esperadas (Index):
    A (0): Código
    B (1): Correo
    C (2): Nombre
    D (3): Esquema
    E (4): Cuenta
    F (5): Apellido (Se usará concatenado con Nombre si es necesario, o visual)
    """
    if pd is None:
        return [], "Pandas no instalado."
    
    try:
        # Leer sin cabecera asumiendo que la data empieza en row 1 (0-indexed) o row 2
        # Frecuentemente tienen encabezados. Intentaremos detectar.
        # Por seguridad leemos header=0.
        df = pd.read_excel(filepath, header=None)
        
        # Opcional: Si la primera fila parece texto de encabezado (ej: "CODIGO"), saltarla.
        # Una heurística simple: si la columna 0 de la row 0 es "CODIGO" o "CÓDIGO"
        if isinstance(df.iloc[0,0], str) and "CODIGO" in df.iloc[0,0].upper():
            df = pd.read_excel(filepath, header=0)
            # Re-leer con header, ahora las columnas tienen nombres, pero accedemos por iloc para ser agnósticos
        
        agentes = []
        for index, row in df.iterrows():
            # Validar que tenga código
            codigo = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ""
            if not codigo or codigo == "CODIGO": continue # Saltarse headers o vacíos
            
            email = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            nombre = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
            esquema = str(row.iloc[3]).strip().upper() if pd.notna(row.iloc[3]) else ""
            cuenta = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
            if cuenta.endswith('.0'): cuenta = cuenta[:-2] # Fix lectura numeros como float
            
            apellido = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else ""
            
            agentes.append({
                "codigo": codigo,
                "email": email,
                "nombre": f"{nombre} {apellido}".strip(),
                "esquema": esquema,
                "cuenta": cuenta
            })
            
        return agentes, None
        
    except Exception as e:
        return [], str(e)

def get_latest_balance_file(root_dir):
    """
    Busca en la carpeta 'Cuentas' el archivo más reciente y retorna un diccionario {Cuenta: Saldo}.
    """
    cuentas_dir = os.path.join(root_dir, "Cuentas")
    if not os.path.exists(cuentas_dir):
        return {}, "No existe carpeta 'Cuentas'", None

    files = [os.path.join(cuentas_dir, f) for f in os.listdir(cuentas_dir) 
             if os.path.isfile(os.path.join(cuentas_dir, f)) and f.lower().endswith(('.xlsx', '.xls'))]
    
    if not files:
        return {}, "No hay archivos en 'Cuentas'", None

    # Ordenar por fecha extraida del nombre, fallback a mtime
    # Tupla de orden: (fecha_nombre, fecha_mtime)
    def file_sort_key(fpath):
        fname = os.path.basename(fpath)
        d = _extract_date_from_name(fname)
        mtime = os.path.getmtime(fpath)
        # Si d existe, usamos timestamp de d. Si no, 0. El segundo criterio es mtime.
        ts_d = d.timestamp() if d else 0
        return (ts_d, mtime)

    latest_file = max(files, key=file_sort_key)
    
    # Debug info
    # print(f"Latest Balance File: {latest_file}")
    
    if pd is None: return {}, "Pandas no instalado", latest_file

    try:
        # Leer archivo de saldos
        # Se asume estructura: Col B=Cuenta, Col J=Saldo (Index 1 y 9) según script original
        df = pd.read_excel(latest_file)
        
        saldos = {}
        # Iterar. Es arriesgado confiar en índices fijos si el formato cambia, pero es lo que tenemos.
        # Buscamos columnas por índice.
        # Col B es index 1, Col J es index 9.
        
        for index, row in df.iterrows():
             # Validación básica de fila
             if len(row) < 10: continue

             cuenta_raw = row.iloc[1]
             saldo_raw = row.iloc[9]
             
             if pd.notna(cuenta_raw) and pd.notna(saldo_raw):
                 cuenta = str(cuenta_raw).strip()
                 if cuenta.endswith('.0'): cuenta = cuenta[:-2]
                 
                 try:
                     val = float(str(saldo_raw).replace(',', '').replace('$', '').strip())
                     saldos[cuenta] = val
                 except:
                     continue
                     
        return saldos, None, latest_file

    except Exception as e:
        return {}, str(e), latest_file

def calculate_debt_for_agent(root_dir, agent, start_date=None, end_date=None):
    """
    Calcula la deuda total de un agente escaneando su carpeta de esquema y TIE.
    start_date: datetime (solo sumar archivos con fecha >= start_date). Si es None, hoy.
    end_date: datetime (solo sumar archivos con fecha <= end_date). Si es None, sin limite superior.
    """
    if start_date is None:
        start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    
    # Asegurar horas 0 para comparacion limpia
    start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    if end_date:
        end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)

    total_debt = 0.0
    details = [] # Lista de {archivo, valor}

    folders_to_check = []
    
    # 1. Esquema Principal
    if agent["esquema"]:
        folders_to_check.append((agent["esquema"], False)) # (NombreCarpeta, EsTIE)
    
    # 2. TIE
    folders_to_check.append(("TIE", True))
    
    for folder_name, is_tie in folders_to_check:
        folder_path = os.path.join(root_dir, folder_name)
        if not os.path.exists(folder_path):
            continue
            
        # Listar archivos
        for fname in os.listdir(folder_path):
            if not fname.lower().endswith(('.xlsx', '.xls')): continue
            
            # Extraer fecha del nombre para filtrar
            try:
                file_date = _extract_date_from_name(fname)
                if not file_date:
                    # Fallback: file mtime
                    file_date = datetime.fromtimestamp(os.path.getmtime(os.path.join(folder_path, fname)))
                
                # Normalizar a media noche para comparacion
                file_date = file_date.replace(hour=0, minute=0, second=0, microsecond=0)
                
                # COMPROBACION DE RANGO
                if file_date < start_date:
                    continue
                if end_date and file_date > end_date:
                    continue

                # Procesar archivo
                fpath = os.path.join(folder_path, fname)
                df = pd.read_excel(fpath) # Header 0 por defecto
                
                # Buscar el Agente (Col 0 por defecto suele ser Código, pero en TIE raw puede ser Col 1)
                
                # Iteramos buscando el código en col 0 y col 1
                found_val = 0.0
                found = False

                col_idx_found = -1
                
                # Check Col 0
                if df.shape[1] > 0:
                     col0 = df.iloc[:, 0].astype(str).str.strip().str.upper()
                     match = df[col0 == agent["codigo"]]
                     if not match.empty:
                         col_idx_found = 0
                
                # Check Col 1 (if not found in 0)
                if col_idx_found == -1 and df.shape[1] > 1:
                     col1 = df.iloc[:, 1].astype(str).str.strip().str.upper()
                     match = df[col1 == agent["codigo"]]
                     if not match.empty:
                         col_idx_found = 1

                if col_idx_found != -1:
                    # Encontrado
                    if is_tie:
                        # TIE: Valor en columna ? 
                        # Si Code en 0 -> Val en 2 (Cleaned)
                        # Si Code en 1 -> Val en 3 (Raw)
                        # Delta es +2
                        val_idx = col_idx_found + 2
                        try:
                            if df.shape[1] > val_idx:
                                val = match.iloc[0, val_idx]
                                found_val = float(val) if pd.notna(val) else 0.0
                                found = True
                        except:
                            pass
                    else:
                        # Esquema normal: Sumar columnas desde la 3 en adelante?
                        # Original: Desde col 3 (D).
                        # Validar si esto cambia si el codigo está en otra col
                        # Asumimos estructura fija para esquemas, solo TIE varia
                        row_vals = match.iloc[0, 3:] # Desde col D
                        found_val = pd.to_numeric(row_vals, errors='coerce').sum()
                        found = True
                
                if found and abs(found_val) > 0.01:
                    total_debt += found_val
                    details.append(f"{fname} ({found_val:,.2f})")

            except Exception as e:
                print(f"Error leyendo {fname}: {e}")
                # details.append(f"[ERROR] {fname}: {str(e)}") # Optional debug
                continue
                
    return total_debt, details

def _extract_date_from_name(filename):
    """
    Intenta extraer fecha de strings como 'GARANTIA 23ENE-2026.xlsx', '23ENE2026', o '2026-02-06'.
    Basado en lógica de Script_automatizacion.txt.
    """
    meses_map = {
        "JAN":0,"FEB":1,"MAR":2,"APR":3,"MAY":4,"JUN":5,"JUL":6,"AUG":7,"SEP":8,"OCT":9,"NOV":10,"DEC":11,
        "ENE":0,"ABR":3,"AGO":7,"DIC":11
    }
    
    filename = filename.upper()
    try:
        import re
        # 1. Regex del Script: (d{1,2})(JAN|FEB...)[- ]?(d{4})
        # Cubre 23ENE2026, 23-ENE-2026, 23 ENE 2026
        meses_regex = "|".join(meses_map.keys())
        pattern1 = r'(\d{1,2})(' + meses_regex + r')[- ]?(\d{4})'
        m1 = re.search(pattern1, filename)
        if m1:
            day = int(m1.group(1))
            month_idx = meses_map[m1.group(2)]
            year = int(m1.group(3))
            return datetime(year, month_idx + 1, day)

        # 2. ISO YYYY-MM-DD o YYYY/MM/DD
        m2 = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', filename)
        if m2:
             return datetime(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))

        # 3. DD-MM-YYYY o DD/MM/YYYY
        m3 = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', filename)
        if m3:
             return datetime(int(m3.group(3)), int(m3.group(2)), int(m3.group(1)))

    except:
        pass
    return None 

def download_scheme_range(start_date, end_date, scheme_name, root_dir, max_workers=20, callback_log=None):
    """
    Descarga archivos para un esquema en un rango de fechas.
    callback_log: Función opcional para recibir mensajes de log (str).
    Retorna: (archivos_descargados, total_dias)
    """
    if scheme_name not in ESQUEMAS:
        if callback_log: callback_log(f"[ERROR] Esquema '{scheme_name}' no existe.")
        return 0, 0
        
    config = ESQUEMAS[scheme_name]
    files_to_try = config["archivos"]
    
    # Crear carpeta
    scheme_folder = os.path.join(root_dir, scheme_name)
    if not os.path.exists(scheme_folder):
        try:
            os.makedirs(scheme_folder)
            if callback_log: callback_log(f"Creada carpeta: {scheme_folder}")
        except OSError as e:
            if callback_log: callback_log(f"[ERROR] No se pudo crear carpeta: {e}")
            return 0, 0
            
    if callback_log: callback_log(f"--- Iniciando Descarga Automática: {scheme_name} ---")
    if callback_log: callback_log(f"Rango: {start_date.strftime('%Y-%m-%d')} a {end_date.strftime('%Y-%m-%d')}")
    
    tasks = []
    versions = ["", "_V2"]
    extensions = [".xlsx", ".XLSX", ".xls", ".XLS"]
    
    current_date = start_date
    delta = timedelta(days=1)
    days_count = 0
    
    while current_date <= end_date:
        days_count += 1
        for file_base in files_to_try:
            variations = [file_base, file_base + " ", file_base.replace(" ", "  ")]
            variations = list(set(variations))
            
            for variant in variations:
                for ver in versions:
                    for ext in extensions:
                        url, filename = get_xm_url(variant, current_date, esquema_nombre=scheme_name, version_suffix=ver, extension=ext)
                        tasks.append((url, filename, scheme_folder, scheme_name))
        current_date += delta
        
    found_count = 0
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_task = {executor.submit(_download_worker_wrapper, *task): task for task in tasks}
        
        for future in concurrent.futures.as_completed(future_to_task):
            try:
                success, msg = future.result()
                if success:
                    found_count += 1
                    if callback_log and msg: callback_log(msg)
                elif msg and "[ERROR]" in msg and callback_log:
                     callback_log(msg)
            except Exception as e:
                if callback_log: callback_log(f"[EXCEPTION] {e}")
                
    if callback_log: callback_log(f"Finalizado {scheme_name}: {found_count} archivos.")
    return found_count, days_count

def _download_worker_wrapper(url, filename, scheme_folder, scheme):
    """Helper interno para procesar descarga y limpieza (logic from GUI)"""
    try:
        success = download_file(url, filename, scheme_folder)
        if success:
            msg = f"[OK] Descargado: {filename}"
            if scheme == "TIE":
                full_path = os.path.join(scheme_folder, filename)
                new_path, error_msg = clean_tie_file(full_path)
                if new_path:
                    new_filename = os.path.basename(new_path)
                    if new_filename != filename:
                         msg += f" -> Limpiado: {new_filename}"
                    else:
                         msg += " -> Limpiado."
                else:
                    msg += f" -> [WARN] Error Limpiando TIE: {error_msg}"
            return True, msg
        return False, None
    except Exception as e:
        return False, f"[ERROR] {filename}: {str(e)}"

if __name__ == "__main__":
    pass
