import sys
print(f"Python: {sys.version}")

try:
    import pandas as pd
    print(f"Pandas: {pd.__version__}")
except ImportError as e:
    print(f"Pandas Error: {e}")

try:
    import openpyxl
    print(f"OpenPyXL: {openpyxl.__version__}")
except ImportError as e:
    print(f"OpenPyXL Error: {e}")

try:
    import xlrd
    print(f"xlrd: {xlrd.__version__}")
except ImportError as e:
    print(f"xlrd Error: {e}")
