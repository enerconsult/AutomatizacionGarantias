
import os
import sys
from datetime import datetime, timedelta
import time

# Ensure we can import the module even if run from a different CWD
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

try:
    import download_xm_file
except ImportError:
    print("Error: No se encontro download_xm_file.py")
    sys.exit(1)

def main():
    # Configuración de fechas
    start_date = datetime.now()
    end_date = start_date + timedelta(days=15)
    
    # Carpeta raíz (relativa al script o absoluta)
    root_dir = os.path.join(current_dir, "Garantías")
    
    print(f"=== INICIANDO EJECUCIÓN AUTOMÁTICA: {datetime.now()} ===")
    print(f"Rango: {start_date.strftime('%Y-%m-%d')} - {end_date.strftime('%Y-%m-%d')}")
    print(f"Carpeta: {root_dir}")
    
    start_time = time.time()
    
    # Iterar todos los esquemas
    total_files = 0
    
    for scheme in download_xm_file.ESQUEMAS.keys():
        print(f"\n>> Procesando Esquema: {scheme}...")
        try:
            # Usamos una lambda para imprimir logs con timestamp simple
            log_func = lambda msg: print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
            
            count, days = download_xm_file.download_scheme_range(
                start_date, 
                end_date, 
                scheme, 
                root_dir, 
                max_workers=10, # Un poco más conservador para background
                callback_log=log_func
            )
            total_files += count
        except Exception as e:
            print(f"[ERROR CRÍTICO] Falló esquema {scheme}: {e}")

    elapsed = time.time() - start_time
    print(f"\n=== PROCESO TERMINADO ===")
    print(f"Total archivos descargados: {total_files}")
    print(f"Tiempo total: {elapsed:.2f} segundos")
    
    # Opcional: Pausa breve si se ejecuta por consola para ver resultado
    # time.sleep(5)

if __name__ == "__main__":
    main()
