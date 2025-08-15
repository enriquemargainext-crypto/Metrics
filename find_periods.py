import pandas as pd
from pathlib import Path

excel_path = Path("../202506_Financials_by_Country.xlsx")
sheet_name = "Mexico Consolidated"

df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

print("Buscando la fila con periodos/fechas:")
print("=" * 50)

# Buscar en las primeras 15 filas para encontrar fechas
for row_idx in range(15):
    row_data = df.iloc[row_idx, 3:15]  # Columnas D a O
    
    # Verificar si hay fechas válidas
    dates_found = []
    for col_idx, cell in enumerate(row_data):
        try:
            if pd.notna(cell):
                # Intentar convertir a fecha
                date_val = pd.to_datetime(cell, errors='coerce')
                if pd.notna(date_val):
                    dates_found.append((col_idx + 3, cell, date_val))
        except:
            pass
    
    if dates_found:
        print(f"Fila {row_idx}: Fechas encontradas:")
        for col_idx, original, parsed in dates_found:
            print(f"  Columna {col_idx}: '{original}' → {parsed}")
        print()

# También verificar las etiquetas de las filas 9 y 10
print("Contenido de filas 9 y 10:")
print(f"Fila 9 (Period): {df.iloc[9, :10].tolist()}")
print(f"Fila 10: {df.iloc[10, :15].tolist()}")

# Buscar directamente donde dice "Period"
period_row = None
for row_idx in range(20):
    for col_idx in range(5):
        cell = df.iloc[row_idx, col_idx]
        if pd.notna(cell) and str(cell).strip().lower() == "period":
            print(f"\n'Period' encontrado en fila {row_idx}, columna {col_idx}")
            # Mostrar las siguientes celdas en esa fila
            following_cells = df.iloc[row_idx, col_idx+1:col_idx+12]
            print(f"Siguientes celdas: {following_cells.tolist()}")
            period_row = row_idx
            break
