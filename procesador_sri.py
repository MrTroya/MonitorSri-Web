#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Procesador SRI
This script replicates the internal Power Query ETL logic from 'Procesador SRI.xlsx'.
It processes the vehicle records for the years 2023, 2024, 2025, and 2026,
cleans the columns, filters out motorcycles, sorts the records chronologically,
and deduplicates the data based on 'CODIGO DE VEHICULO' (keeping only the earliest record).

Usage:
    python procesador_sri.py --input-dir "/path/to/csv/dir" --output-file "output.csv"
"""

import os
import re
import argparse
import time
import urllib.request
import ssl
import pandas as pd

# Spanish month abbreviation dictionary to translate month strings to integers
SPANISH_MONTHS = {
    'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'ago': 8, 'sep': 9, 'sept': 9, 'oct': 10, 'nov': 11, 'dic': 12
}

def parse_spanish_date(val):
    """
    Parses Spanish dates in 'd-mmm-yy' format (e.g. 2-nov-23) to pandas Timestamps.
    Fails back to standard pandas datetime parser if format doesn't match.
    """
    if pd.isna(val) or not isinstance(val, str):
        return pd.NaT
    val = val.strip().lower()
    if not val:
        return pd.NaT
    
    # Check for DD-MMM-YY format (using Spanish month abbreviations)
    m = re.match(r'^(\d+)-([a-z]+)-(\d+)$', val)
    if m:
        day = int(m.group(1))
        month_str = m.group(2)
        year_short = int(m.group(3))
        month = SPANISH_MONTHS.get(month_str)
        if month:
            year = 2000 + year_short
            try:
                return pd.Timestamp(year=year, month=month, day=day)
            except ValueError:
                return pd.NaT
                
    # Fallback to pandas automatic date conversion (with dayfirst enabled)
    return pd.to_datetime(val, errors='coerce', dayfirst=True)

def ensure_csv_file(year, target_dir):
    """
    Checks if the CSV file for the given year exists locally.
    If not, downloads it dynamically from the official SRI download page.
    """
    filename = f"SRI_Vehiculos_Nuevos_{year}.csv"
    filepath = os.path.join(target_dir, filename)
    if os.path.exists(filepath):
        print(f"Archivo encontrado localmente: {filepath}")
        return filepath
        
    url = f"https://descargas.sri.gob.ec/download/datosAbiertos/{filename}"
    print(f"Archivo {filename} no encontrado en {target_dir}. Descargándolo desde {url}...")
    
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        os.makedirs(target_dir, exist_ok=True)
        with urllib.request.urlopen(req, context=ctx, timeout=90) as response:
            with open(filepath, 'wb') as out_file:
                out_file.write(response.read())
        print(f"Descargado exitosamente: {filepath}")
    except Exception as e:
        # Intento de reintento simple para redes inestables
        print(f"Error al descargar {filename} ({e}). Reintentando en 5 segundos...")
        time.sleep(5)
        try:
            with urllib.request.urlopen(req, context=ctx, timeout=90) as response:
                with open(filepath, 'wb') as out_file:
                    out_file.write(response.read())
            print(f"Descargado exitosamente tras reintento: {filepath}")
        except Exception as retry_err:
            raise FileNotFoundError(f"No se pudo descargar {filename} desde {url}: {retry_err}")
            
    return filepath

def load_and_clean_2023(source_dir):
    """Loads and cleans SRI_Vehiculos_Nuevos_2023.csv"""
    filepath = ensure_csv_file(2023, source_dir)
    df = pd.read_csv(filepath, sep=';', encoding='latin-1')
    
    cols_to_keep = [
        "CATEGORÍA", "CODIGO DE VEHICULO", "TIPO TRANSACCIÓN", "MARCA", "MODELO",
        "CLASE", "SUB CLASE", "TIPO", "FECHA PROCESO (MM/DD/AA)", "FECHA COMPRA (DD/MM/AA)"
    ]
    df = df[cols_to_keep].copy()
    df = df[df["CLASE"] != "MOTOCICLETA"]
    df = df.rename(columns={
        "FECHA PROCESO (MM/DD/AA)": "FECHA PROCESO",
        "FECHA COMPRA (DD/MM/AA)": "FECHA COMPRA"
    })
    
    df["FECHA PROCESO"] = df["FECHA PROCESO"].apply(parse_spanish_date)
    df["FECHA COMPRA"] = df["FECHA COMPRA"].apply(parse_spanish_date)
    return df

def load_and_clean_2024(source_dir):
    """Loads and cleans SRI_Vehiculos_Nuevos_2024.csv"""
    filepath = ensure_csv_file(2024, source_dir)
    df = pd.read_csv(filepath, sep=';', encoding='latin-1')
    
    cols_to_keep = [
        "CATEGORÍA", "CODIGO DE VEHICULO", "TIPO TRANSACCIÓN", "MARCA", "MODELO",
        "CLASE", "SUB CLASE", "TIPO", "FECHA PROCESO (MM/DD/AA)", "FECHA COMPRA (DD/MM/AA)"
    ]
    df = df[cols_to_keep].copy()
    df = df[df["CLASE"] != "MOTOCICLETA"]
    df = df.rename(columns={
        "FECHA PROCESO (MM/DD/AA)": "FECHA PROCESO",
        "FECHA COMPRA (DD/MM/AA)": "FECHA COMPRA"
    })
    
    df["FECHA PROCESO"] = pd.to_datetime(df["FECHA PROCESO"], format='%d/%m/%Y', errors='coerce')
    df["FECHA COMPRA"] = pd.to_datetime(df["FECHA COMPRA"], format='%d/%m/%Y', errors='coerce')
    return df

def load_and_clean_2025(source_dir):
    """Loads and cleans SRI_Vehiculos_Nuevos_2025.csv"""
    filepath = ensure_csv_file(2025, source_dir)
    df = pd.read_csv(filepath, sep=';', encoding='latin-1')
    
    cols_to_keep = [
        "CATEGORÍA", "CÓDIGO DE VEHÍCULO", "TIPO TRANSACCIÓN", "MARCA", "MODELO",
        "CLASE", "SUB CLASE", "TIPO", "FECHA PROCESO (DD/MM/AAAA)", "FECHA COMPRA (DD/MM/AAAA)"
    ]
    df = df[cols_to_keep].copy()
    df = df[df["CLASE"] != "MOTOCICLETA"]
    df = df.rename(columns={
        "CÓDIGO DE VEHÍCULO": "CODIGO DE VEHICULO",
        "FECHA PROCESO (DD/MM/AAAA)": "FECHA PROCESO",
        "FECHA COMPRA (DD/MM/AAAA)": "FECHA COMPRA"
    })
    
    df["FECHA PROCESO"] = pd.to_datetime(df["FECHA PROCESO"], format='%d/%m/%Y', errors='coerce')
    df["FECHA COMPRA"] = pd.to_datetime(df["FECHA COMPRA"], format='%d/%m/%Y', errors='coerce')
    return df

def load_and_clean_2026(source_dir):
    """Loads and cleans SRI_Vehiculos_Nuevos_2026.csv"""
    filepath = ensure_csv_file(2026, source_dir)
    df = pd.read_csv(filepath, sep=';', encoding='latin-1')
    
    cols_to_keep = [
        "CATEGORÍA", "CÓDIGO DE VEHÍCULO", "TIPO TRANSACCIÓN", "MARCA", "MODELO",
        "CLASE", "SUB CLASE", "TIPO", "FECHA PROCESO (DD/MM/AAAA)", "FECHA COMPRA (DD/MM/AAAA)"
    ]
    df = df[cols_to_keep].copy()
    df = df[df["CLASE"] != "MOTOCICLETA"]
    df = df.rename(columns={
        "CÓDIGO DE VEHÍCULO": "CODIGO DE VEHICULO",
        "FECHA PROCESO (DD/MM/AAAA)": "FECHA PROCESO",
        "FECHA COMPRA (DD/MM/AAAA)": "FECHA COMPRA"
    })
    
    df["FECHA PROCESO"] = pd.to_datetime(df["FECHA PROCESO"], format='%d/%m/%Y', errors='coerce')
    df["FECHA COMPRA"] = pd.to_datetime(df["FECHA COMPRA"], format='%d/%m/%Y', errors='coerce')
    return df

def main():
    parser = argparse.ArgumentParser(description="Procesa datos vehiculares SRI replicando la lógica de Power Query.")
    parser.add_argument(
        "--input-dir",
        default=os.path.dirname(os.path.abspath(__file__)),
        help="Directorio que contiene los archivos CSV origen (por defecto el directorio del script)."
    )
    parser.add_argument(
        "--output-file",
        default="Procesador_SRI_Result.csv",
        help="Nombre base del archivo de salida (se guardará como .csv y .xlsx)."
    )
    
    args = parser.parse_args()
    
    start_time = time.time()
    print(f"Iniciando procesamiento en el directorio: {args.input_dir}")
    
    # Step 1: Load and clean each dataset
    try:
        print("Cargando y procesando datos 2023...")
        df_2023 = load_and_clean_2023(args.input_dir)
        print(f"  Registros 2023 (sin motocicletas): {len(df_2023)}")
        
        print("Cargando y procesando datos 2024...")
        df_2024 = load_and_clean_2024(args.input_dir)
        print(f"  Registros 2024 (sin motocicletas): {len(df_2024)}")
        
        print("Cargando y procesando datos 2025...")
        df_2025 = load_and_clean_2025(args.input_dir)
        print(f"  Registros 2025 (sin motocicletas): {len(df_2025)}")
        
        print("Cargando y procesando datos 2026...")
        df_2026 = load_and_clean_2026(args.input_dir)
        print(f"  Registros 2026 (sin motocicletas): {len(df_2026)}")
        
    except FileNotFoundError as e:
        print(f"Error: {e}")
        print("Asegúrate de que los archivos CSV existan o que haya una conexión a Internet para descargarlos.")
        return
        
    # Step 2: Combine datasets
    print("Combinando tablas de todos los años...")
    combined = pd.concat([df_2023, df_2024, df_2025, df_2026], ignore_index=True)
    print(f"  Total de registros combinados: {len(combined)}")
    
    # Step 3: Sort by FECHA PROCESO ascending
    print("Ordenando filas cronológicamente por FECHA PROCESO...")
    combined = combined.sort_values(by="FECHA PROCESO", ascending=True).reset_index(drop=True)
    
    # Step 4: Add sequential Index (equivalent to Table.AddIndexColumn starting from 0)
    combined["Index"] = combined.index
    
    # Step 5: Group and find Min Index for each CODIGO DE VEHICULO
    print("Agrupando y determinando el índice mínimo por código de vehículo (deduplicación)...")
    min_idx = combined.groupby("CODIGO DE VEHICULO")["Index"].min().reset_index()
    min_idx = min_idx.rename(columns={"Index": "Cod Min Vehiculo.Min"})
    
    # Step 6: Merge the min index back and filter
    print("Fusionando índices y aplicando filtro de unicidad...")
    merged = pd.merge(combined, min_idx, on="CODIGO DE VEHICULO", how="left")
    merged["Valid Usar"] = merged["Cod Min Vehiculo.Min"] == merged["Index"]
    
    final_df = merged[merged["Valid Usar"] == True].copy()
    print(f"  Total de registros finales deduplicados: {len(final_df)}")
    
    # Keep Datetime copy for calculating the latest month summary before string formatting
    final_df["FECHA_PROCESO_DT"] = pd.to_datetime(final_df["FECHA PROCESO"])
    
    # Step 7: Reorder columns to match the output Table schema
    cols_order = [
        "Index", "CATEGORÍA", "CODIGO DE VEHICULO", "TIPO TRANSACCIÓN", "MARCA", "MODELO",
        "CLASE", "SUB CLASE", "TIPO", "FECHA PROCESO", "FECHA COMPRA", "Cod Min Vehiculo.Min", "Valid Usar"
    ]
    
    # Convert dates to string YYYY-MM-DD for final output formatting
    final_df["FECHA PROCESO"] = final_df["FECHA PROCESO"].dt.strftime('%Y-%m-%d')
    final_df["FECHA COMPRA"] = final_df["FECHA COMPRA"].dt.strftime('%Y-%m-%d')
    
    # Convert logical to integer 1/0 for Valid Usar to match Excel representation
    final_df["Valid Usar"] = final_df["Valid Usar"].astype(int)
    
    # Final slice to order
    output_df = final_df[cols_order].copy()
    
    # Step 8: Save outputs
    base_out = args.output_file
    if base_out.endswith(".csv") or base_out.endswith(".xlsx"):
        base_out = os.path.splitext(base_out)[0]
        
    csv_out = base_out + ".csv"
    xlsx_out = base_out + ".xlsx"
    
    print(f"Guardando resultado CSV en: {csv_out}...")
    output_df.to_csv(csv_out, sep=';', index=False, encoding='latin-1')
    
    print(f"Guardando resultado Excel en: {xlsx_out}...")
    try:
        output_df.to_excel(xlsx_out, index=False)
        print("  ¡Archivo Excel guardado con éxito!")
    except Exception as e:
        print(f"  Error al guardar en Excel: {e}")
        
    # Step 9: Dynamic Console summary for the LATEST month in dataset
    max_date = final_df["FECHA_PROCESO_DT"].max()
    max_year = max_date.year
    max_month = max_date.month
    
    latest_month_df = final_df[
        (final_df["FECHA_PROCESO_DT"].dt.year == max_year) & 
        (final_df["FECHA_PROCESO_DT"].dt.month == max_month)
    ]
    total_latest_sales = len(latest_month_df)
    
    # Group by MARCA and calculate sales
    brand_sales = latest_month_df.groupby("MARCA").size().reset_index(name="Ventas")
    top_10 = brand_sales.sort_values(by="Ventas", ascending=False).head(10).copy()
    top_10["MarketShare"] = (top_10["Ventas"] / total_latest_sales * 100).map("{:.2f}%".format)
    
    month_names = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
        7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
    }
    month_name = month_names.get(max_month, str(max_month))
    
    print(f"\n==================================================")
    print(f"RESUMEN DEL ÚLTIMO MES DISPONIBLE: {month_name.upper()} {max_year}")
    print(f"Total de ventas del mes: {total_latest_sales:,}")
    print(f"==================================================")
    print("| Pos | Marca | Ventas | Market Share |")
    print("|---|---|---|---|")
    for idx, row in enumerate(top_10.itertuples(), 1):
        print(f"| {idx} | {row.MARCA} | {row.Ventas:,} | {row.MarketShare} |")
    print(f"==================================================")
    
    elapsed = time.time() - start_time
    print(f"Procesamiento finalizado con éxito en {elapsed:.2f} segundos.")

if __name__ == "__main__":
    main()
