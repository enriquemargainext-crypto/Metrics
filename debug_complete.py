import pandas as pd
from pathlib import Path

def read_sheet_as_dataframe_debug(excel_path, sheet_name):
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, engine="openpyxl")

    print("Buscando fila con fechas válidas...")
    
    # Buscar en las primeras 15 filas
    header_row_idx = None
    start_col_idx = None
    
    for row_idx in range(15):
        row = df.iloc[row_idx, :]
        for col_idx, cell in enumerate(row):
            if pd.notna(cell):
                try:
                    date_val = pd.to_datetime(cell, errors='coerce')
                    if pd.notna(date_val) and date_val.year > 2020:
                        print(f"Fecha encontrada en fila {row_idx}, columna {col_idx}: {cell} → {date_val}")
                        if header_row_idx is None:
                            header_row_idx = row_idx
                            start_col_idx = col_idx
                        break
                except:
                    pass
        if header_row_idx is not None:
            break
    
    if header_row_idx is None:
        print("❌ No se encontraron fechas válidas")
        return df, []
    
    print(f"Usando fila {header_row_idx}, comenzando en columna {start_col_idx}")
    
    # Extraer períodos de esa fila desde esa columna
    raw_periods = df.iloc[header_row_idx, start_col_idx:]
    print(f"Raw periods: {raw_periods.tolist()[:10]}")
    
    # Filtramos columnas que realmente tienen fecha
    raw_periods = raw_periods.dropna()
    print(f"Periods después de dropna: {raw_periods.tolist()}")
    
    periods = pd.to_datetime(raw_periods, errors="coerce").dropna()
    print(f"Periods después de to_datetime: {periods.tolist()}")
    
    periods = list(pd.DatetimeIndex(periods).to_period("M").to_timestamp("M"))
    print(f"Periods finales: {periods}")

    return df, periods, start_col_idx

def extract_metric_debug(df, periods, metric_label, search_column, data_start_col):
    start_col_idx = data_start_col
    
    print(f"\n--- Extrayendo '{metric_label}' de columna {search_column} ---")
    
    # Buscar la métrica
    labels_col = df.iloc[:, search_column]
    matches = labels_col[labels_col == metric_label]
    
    if matches.empty:
        print(f"❌ No encontrado: '{metric_label}'")
        return None
    
    row_idx = matches.index[0]
    print(f"✅ Encontrado en fila {row_idx}")
    
    # Extraer valores
    row_values = df.iloc[row_idx, start_col_idx : start_col_idx + len(periods)]
    print(f"Raw values: {row_values.tolist()}")
    
    values = pd.to_numeric(row_values, errors="coerce").astype(float)
    print(f"Numeric values: {values.tolist()}")
    
    series = pd.Series(values.values, index=pd.DatetimeIndex(periods), name=metric_label)
    print(f"Series final: {series}")
    
    return series

# Test completo
excel_path = Path("../202506_Financials_by_Country.xlsx")
sheet_name = "Mexico Consolidated"

print("=== DEBUG COMPLETO ===")
df, periods, data_start_col = read_sheet_as_dataframe_debug(excel_path, sheet_name)

print(f"\nPeriodos extraídos: {len(periods)}")
print(f"Columna de inicio de datos: {data_start_col}")

if len(periods) > 0:
    # Probar extracción de métricas
    test_metrics = [
        ("Cars Sold - Delivered", 2),
        ("Total Net Revenues", 3),
    ]
    
    for metric, col in test_metrics:
        series = extract_metric_debug(df, periods, metric, col, data_start_col)
else:
    print("❌ No se pudieron extraer periodos")
