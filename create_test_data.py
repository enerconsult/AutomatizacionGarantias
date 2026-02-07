
import pandas as pd
import os
from datetime import datetime

# Create dummy Maestro
maestro_data = [
    ["AG001", "test@example.com", "Agente 1", "Mensual", "12345", "Apellido1"],
    ["AG002", "test2@example.com", "Agente 2", "Semanal", "67890", "Apellido2"]
]
df_maestro = pd.DataFrame(maestro_data)
df_maestro.to_excel("test_maestro.xlsx", index=False, header=False)
print("Created test_maestro.xlsx")

# Create dummy Balance file
os.makedirs("Descargas_XM/Cuentas", exist_ok=True)
balance_data = {
    "ColA": ["X"] * 2,
    "Cuenta": ["12345", "67890"],
    "ColC": ["X"] * 2,
    "ColD": ["X"] * 2,
    "ColE": ["X"] * 2,
    "ColF": ["X"] * 2,
    "ColG": ["X"] * 2,
    "ColH": ["X"] * 2,
    "ColI": ["X"] * 2,
    "Saldo": [1000000, 500000]
}
df_balance = pd.DataFrame(balance_data)
# Ensure columns represent indices correctly (B=1, J=9)
# Current: 0, 1, 2, ..., 9
df_balance.to_excel("Descargas_XM/Cuentas/Saldo test.xlsx", index=False)
print("Created Descargas_XM/Cuentas/Saldo test.xlsx")

# Create dummy Debt files
os.makedirs("Descargas_XM/Mensual", exist_ok=True)
# Code in Col A (0), Value in Col D (3)
debt_data = [
    ["AG001", "X", "X", 120000, "X"],
    ["OTHER", "X", "X", 0, "X"]
]
df_debt = pd.DataFrame(debt_data)
df_debt.to_excel("Descargas_XM/Mensual/GARANTIA MENSUAL 07FEB-2026.xlsx", index=False)
print("Created Descargas_XM/Mensual/GARANTIA MENSUAL 07FEB-2026.xlsx")
