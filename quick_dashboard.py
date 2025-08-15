import pandas as pd
from pathlib import Path
import argparse

def quick_extract(excel_path, sheet_name="Mexico Consolidated"):
    """Extracción súper rápida - solo métricas esenciales"""
    
    print(f"📊 Extrayendo datos de {sheet_name}...")
    
    # Leer solo las columnas necesarias (A:P para ser conservadores)
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, 
                      usecols="A:P", engine="openpyxl")
    
    # Extraer períodos (fila 10, desde columna J)
    periods_row = df.iloc[9, 9:]  # Fila 10, desde columna J
    periods = pd.to_datetime(periods_row.dropna(), errors="coerce").dropna()
    periods = periods[-3:]  # Solo últimos 3 meses para máxima velocidad
    
    print(f"📅 Períodos: {[p.strftime('%Y-%m') for p in periods]}")
    
    # Métricas esenciales solamente
    essential_metrics = {
        "Cars Sold": ("Cars Sold - Delivered", 2),
        "Inventory": ("Inventory BoM", 2), 
        "Revenues": ("Total Net Revenues", 3),
        "EBITDA": ("Adj. EBITDA", 3)
    }
    
    results = []
    
    for metric_name, (label, search_col) in essential_metrics.items():
        print(f"🔍 Buscando: {metric_name}")
        
        # Buscar la etiqueta
        labels_col = df.iloc[:, search_col].astype(str)
        match_idx = None
        
        for i, cell_value in enumerate(labels_col):
            if label in cell_value:
                match_idx = i
                break
        
        if match_idx is not None:
            # Extraer valores (últimos 3 períodos solamente)
            row_data = df.iloc[match_idx, 9:9+len(periods)]
            values = pd.to_numeric(row_data, errors="coerce")
            
            for period, value in zip(periods, values):
                if pd.notna(value):
                    results.append({
                        'metric': metric_name,
                        'period': period.strftime('%Y-%m-%d'),
                        'value': value
                    })
            
            print(f"  ✅ {metric_name}: {len(values)} valores")
        else:
            print(f"  ❌ {metric_name}: No encontrado")
    
    return pd.DataFrame(results)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Dashboard rápido")
    parser.add_argument("--excel_path", default="../202506_Financials_by_Country.xlsx")
    parser.add_argument("--sheet", default="Mexico Consolidated")
    args = parser.parse_args()
    
    excel_path = Path(args.excel_path)
    if not excel_path.exists():
        print(f"❌ Archivo no encontrado: {excel_path}")
        exit(1)
    
    print("🚀 DASHBOARD RÁPIDO - Solo métricas esenciales")
    print("=" * 50)
    
    try:
        df_result = quick_extract(excel_path, args.sheet)
        
        if not df_result.empty:
            print("\n📋 RESULTADOS:")
            print(df_result.to_string(index=False))
            
            # Guardar CSV
            csv_file = f"quick_results_{args.sheet.replace(' ', '_')}.csv"
            df_result.to_csv(csv_file, index=False)
            print(f"\n💾 Guardado en: {csv_file}")
        else:
            print("❌ No se encontraron datos")
            
    except Exception as e:
        print(f"❌ Error: {e}")
