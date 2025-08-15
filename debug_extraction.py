import pandas as pd
from pathlib import Path

# Usar el mismo código de extracción del dashboard
def _normalize_label(text: str) -> str:
    if not isinstance(text, str):
        return ""
    normalized = (
        text.replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u2212", "-")
        .strip()
    )
    return normalized

def extract_metric_series_debug(
    df: pd.DataFrame, periods, metric_label_exact: str, search_column: int = 2
):
    start_col_idx = 3
    
    print(f"\nBuscando '{metric_label_exact}' en columna {search_column}:")
    
    # Normalizamos para evitar diferencias de guiones
    target = _normalize_label(metric_label_exact)
    labels_col = df.iloc[:, search_column].apply(_normalize_label)
    matches = labels_col[labels_col == target]
    
    print(f"  Target normalizado: '{target}'")
    print(f"  Etiquetas encontradas en columna {search_column}:")
    for i, label in enumerate(labels_col[:50]):  # Primeras 50
        if label and target.lower() in label.lower():
            print(f"    Fila {i}: '{label}' {'✅ MATCH' if label == target else '⚠️ SIMILAR'}")
    
    if matches.empty:
        print(f"  ❌ No se encontró coincidencia exacta")
        return None
    
    row_idx = matches.index[0]
    print(f"  ✅ Encontrado en fila {row_idx}")
    
    # Verificar los valores
    row_values = df.iloc[row_idx, start_col_idx:start_col_idx + 10]  # Primeras 10 columnas de datos
    print(f"  Primeros valores: {row_values.tolist()}")
    
    return row_idx

# Probar con datos reales
excel_path = Path("../202506_Financials_by_Country.xlsx")
sheet_name = "Mexico Consolidated"

print(f"Debuggeando extracción de métricas")
print("=" * 50)

df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

# Verificar periodos en fila 5
print("Periodos en fila 5 (índice 4):")
period_row = df.iloc[4, 3:13]
print(period_row.tolist())

print("\n" + "=" * 50)

# Probar algunas métricas específicas
test_metrics = [
    ("Cars Sold - Delivered", 2),
    ("Cars Purchased", 2), 
    ("Inventory BoM", 2),
    ("Total Net Revenues", 3),
    ("Metal Margin (mm)", 3),
    ("PC1 (mm)", 3),
]

for metric, col in test_metrics:
    extract_metric_series_debug(df, [], metric, col)
