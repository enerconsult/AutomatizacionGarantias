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

    # Ordenar por fecha de modificación (o nombre si tiene fecha)
    # Asumimos modificación por simplicidad, o nombre si queremos ser estrictos con la fecha en nombre
    latest_file = max(files, key=os.path.getmtime)
    file_date = datetime.fromtimestamp(os.path.getmtime(latest_file)).strftime("%Y-%m-%d")

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

def calculate_debt_for_agent(root_dir, agent, date_filter=None):
    """
    Calcula la deuda total de un agente escaneando su carpeta de esquema y TIE.
    date_filter: datetime (solo sumar archivos con fecha >= date_filter). Si es None, hoy.
    """
    if date_filter is None:
        date_filter = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

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
            # Reutilizamos lógica simple o regex
            # Aquí haremos un parsing básico.
            try:
                # Intento de heurística de fecha. 
                # Si falla, asumimos que es reciente si la fecha de modificación es reciente?
                # Mejor usar la fecha del archivo si es posible.
                file_date = _extract_date_from_name(fname)
                if not file_date:
                    # Fallback: file mtime
                    file_date = datetime.fromtimestamp(os.path.getmtime(os.path.join(folder_path, fname)))
                
                # Normalizar a media noche
                file_date = file_date.replace(hour=0, minute=0, second=0, microsecond=0)
                
                if file_date < date_filter:
                    continue

                # Procesar archivo
                fpath = os.path.join(folder_path, fname)
                df = pd.read_excel(fpath) # Header 0 por defecto
                
                # Buscar el Agente (Col 0 por defecto suele ser Código)
                # TIE tiene estructura diferente (col 0 borrada, ahora col 0 es codigo?)
                # Asumimos que la columna "Agente" o "Codigo" está en la primera posición disponible
                
                # Iteramos buscando el código
                found_val = 0.0
                found = False

                # Convert to string and upper for matching
                # Check first column
                col0 = df.iloc[:, 0].astype(str).str.strip().str.upper()
                match = df[col0 == agent["codigo"]]
                
                if not match.empty:
                    # Encontrado
                    if is_tie:
                        # TIE: Valor en columna ? (Original script: col 3 -> D)
                        # Al borrar la primera col, se corre.
                        # Asumamos que buscamos un valor numérico en las cols siguientes.
                        # Original: col 3 (D). Si borramos A, ahora es C (2).
                        try:
                            val = match.iloc[0, 2] # Index 2 = Col C
                            found_val = float(val) if pd.notna(val) else 0.0
                            found = True
                        except:
                            pass
                    else:
                        # Esquema normal: Sumar columnas desde la 3 en adelante?
                        # Original script: "for let c = 3; c < matriz[f].length; c++"
                        # Index 3 es Col D.
                        row_vals = match.iloc[0, 3:] # Desde col D
                        found_val = pd.to_numeric(row_vals, errors='coerce').sum()
                        found = True
                
                if found and abs(found_val) > 0.01:
                    total_debt += found_val
                    details.append(f"{fname} ({found_val:,.2f})")

            except Exception as e:
                print(f"Error leyendo {fname}: {e}")
                continue
                
    return total_debt, details

def _extract_date_from_name(filename):
    """ Intenta extraer fecha de strings como 'GARANTIA 23ENE-2026.xlsx' o '2026-02-06' """
    # Diccionario meses
    meses = {"ENE":1, "FEB":2, "MAR":3, "ABR":4, "MAY":5, "JUN":6, 
             "JUL":7, "AGO":8, "SEP":9, "OCT":10, "NOV":11, "DIC":12,
             "JAN":1, "APR":4, "AUG":8, "DEC":12} # Ingles/Español mix
    
    filename = filename.upper()
    try:
        # ISO YYYY-MM-DD
        import re
        iso_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
        if iso_match:
            return datetime(int(iso_match.group(1)), int(iso_match.group(2)), int(iso_match.group(3)))
            
        # Custom 23ENE-2026 (o similar)
        # Buscamos combinaciones DDMMM
        for mes, num in meses.items():
            if mes in filename:
                # Buscar dia antes
                # Regex: (\d{1,2})MES
                day_match = re.search(r'(\d{1,2})' + mes, filename)
                year_match = re.search(r'(\d{4})', filename)
                
                if day_match and year_match:
                    return datetime(int(year_match.group(0)), num, int(day_match.group(1)))
    except:
        return None
    return None

if __name__ == "__main__":
    pass

