import requests
from bs4 import BeautifulSoup
import urllib.parse

def find_xm_file_paths():
    url = "https://www.xm.com.co/administraci%C3%B3n-financiera/garant%C3%ADas-financieras/c%C3%A1lculos-garant%C3%ADas-financieras-mensuales-esquema"
    
    # Use a standard browser User-Agent to avoid 403 Forbidden
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        print(f"Intentando acceder a: {url}")
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        print("\n--- Enlaces encontrados (XLSX, XLS, PDF) ---")
        found = False
        for link in soup.find_all('a', href=True):
            href = link['href']
            # Check for common file extensions
            if href.lower().endswith(('.xlsx', '.xls', '.pdf')):
                full_url = urllib.parse.urljoin(url, href)
                print(f"Archivo: {link.get_text(strip=True)}")
                print(f"Ruta: {full_url}")
                print("-" * 30)
                found = True
                
        if not found:
            print("No se encontraron archivos con extensiones .xlsx, .xls o .pdf directamente en los enlaces.")
            
    except requests.exceptions.RequestException as e:
        print(f"Error al acceder a la página: {e}")

if __name__ == "__main__":
    find_xm_file_paths()
