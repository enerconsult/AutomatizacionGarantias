
import sys
import os
# Add current dir to path to import download_xm_file
sys.path.append(os.getcwd())
import download_xm_file

def test_logic():
    print("Testing read_maestro_file...")
    agentes, err = download_xm_file.read_maestro_file("test_maestro.xlsx")
    if err:
        print(f"Error reading maestro: {err}")
    else:
        print(f"Read {len(agentes)} agents.")
        for a in agentes:
            print(f"  - {a}")

    print("\nTesting get_latest_balance_file...")
    saldos, err, file = download_xm_file.get_latest_balance_file(".") 
    # Note: "." because create_test_data made Descargas_XM relative to Cwd
    
    # Actually get_latest_balance_file takes root_dir and appends "Cuentas"
    # My test data is in Descargas_XM/Cuentas
    # So root_dir should be Descargas_XM
    saldos, err, file = download_xm_file.get_latest_balance_file("Descargas_XM")
    
    if err:
        print(f"Error reading balances: {err}")
    else:
        print(f"Balances found in {file}:")
        print(saldos)

    print("\nTesting calculate_debt_for_agent...")
    if agentes:
        agent = agentes[0] # AG001
        deuda, detalles = download_xm_file.calculate_debt_for_agent("Descargas_XM", agent)
        print(f"Deuda for {agent['codigo']}: {deuda}")
        print(f"Detalles: {detalles}")

if __name__ == "__main__":
    test_logic()
