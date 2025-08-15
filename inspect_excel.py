import pandas as pd

# Inspeccionar el archivo Excel
excel_path = "../202506_Financials_by_Country.xlsx"
sheet_name = "Mexico Consolidated"

print(f"Inspeccionando archivo: {excel_path}")
print(f"Hoja: {sheet_name}")
print("-" * 50)

# Leer sin header
df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

print("Primeras 120 filas - Columnas B, C, D:")
for i in range(120):
    col_b = df.iloc[i, 1] if i < len(df) else ""
    col_c = df.iloc[i, 2] if i < len(df) else ""
    col_d = df.iloc[i, 3] if i < len(df) else ""
    
    if pd.notna(col_b) or pd.notna(col_c) or pd.notna(col_d):
        b_text = str(col_b).strip() if pd.notna(col_b) else ""
        c_text = str(col_c).strip() if pd.notna(col_c) else ""
        d_text = str(col_d).strip() if pd.notna(col_d) else ""
        
        if b_text or c_text or (d_text and not d_text.replace('.','').replace('-','').isdigit()):
            print(f"  Fila {i}: B='{b_text}' | C='{c_text}' | D='{d_text}'")

print("\n" + "-" * 50)
print("Fila 5 (índice 4) - Encabezados de período:")
period_row = df.iloc[4, 3:10]  # Primeras columnas de período
for i, period in enumerate(period_row):
    if pd.notna(period):
        print(f"  Columna {i+3}: '{period}'")
