"""
Robust Parser for 8-Column Trial Balances (Balance de 8 Columnas)
Supports Excel (.xlsx, .xls) and CSV formats with automated header detection and currency sanitization.
"""
import io
import re
import pandas as pd
import numpy as np
from typing import Dict, Any, Tuple, Optional, List

# Column aliases mapping
COLUMN_PATTERNS = {
    "codigo": [r"^c[oó]d", r"^n[o°º]?\s*cta", r"^cuenta$", r"^n_cuenta$", r"^c[oó]digo\s*cuenta", r"^id$"],
    "nombre_cuenta": [r"^nombre", r"^descripci[oó]n", r"^glosa", r"^detalle", r"^cuenta\s*nombre", r"^denominaci[oó]n"],
    "debito": [r"^d[eé]bito", r"^debitos", r"^sumas?\s*d[eé]bito", r"^total\s*d[eé]bitos?"],
    "credito": [r"^cr[eé]dito", r"^creditos", r"^sumas?\s*cr[eé]dito", r"^total\s*cr[eé]ditos?"],
    "deudor": [r"^deudor", r"^saldos?\s*deudor", r"^saldo\s*d$"],
    "acreedor": [r"^acreedor", r"^saldos?\s*acreedor", r"^saldo\s*a$"],
    "activo": [r"^activo", r"^activos", r"^inventario\s*activo"],
    "pasivo": [r"^pasivo", r"^pasivos", r"^inventario\s*pasivo"],
    "perdida": [r"^p[eé]rdida", r"^perdidas", r"^gastos?", r"^resultado\s*p[eé]rdida"],
    "ganancia": [r"^ganancia", r"^ganancias", r"^ingresos?", r"^resultado\s*ganancia"]
}

def clean_financial_number(val: Any) -> float:
    """Converts diverse string financial values (e.g., '$ 1.234.567', '(12,345.50)', '-1.000') into pure float."""
    if pd.isna(val) or val is None or val == "":
        return 0.0
    if isinstance(val, (int, float, np.number)):
        return float(val) if not np.isnan(val) else 0.0
    
    val_str = str(val).strip()
    if not val_str or val_str == "-" or val_str.lower() == "nan":
        return 0.0
    
    # Check for accounting parentheses e.g. (1,500) -> -1500
    is_negative = False
    if val_str.startswith("(") and val_str.endswith(")"):
        is_negative = True
        val_str = val_str[1:-1].strip()
    elif val_str.startswith("-"):
        is_negative = True
        val_str = val_str[1:].strip()
    
    # Remove currency signs and spaces
    val_str = re.sub(r"[\$€£CLPUSD\s]", "", val_str)
    
    # Determine decimal vs thousands separators
    # Case 1: 1.234.567,89 (Chilean/Latin standard)
    if "." in val_str and "," in val_str:
        if val_str.rfind(",") > val_str.rfind("."):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "." in val_str and val_str.count(".") > 1:
        # e.g. 1.234.567
        val_str = val_str.replace(".", "")
    elif "," in val_str and val_str.count(",") > 1:
        # e.g. 1,234,567
        val_str = val_str.replace(",", "")
    elif "," in val_str and len(val_str.split(",")[-1]) != 2:
        # e.g. 1,000 without decimal
        val_str = val_str.replace(",", "")
    elif "," in val_str:
        # e.g. 1234,56
        val_str = val_str.replace(",", ".")
        
    try:
        num = float(val_str)
        return -num if is_negative else num
    except ValueError:
        return 0.0

def match_column(col_name: str) -> Optional[str]:
    """Matches raw column names against standardized financial columns."""
    normalized = str(col_name).strip().lower()
    # Direct check
    for canonical_name, patterns in COLUMN_PATTERNS.items():
        for pat in patterns:
            if re.search(pat, normalized):
                return canonical_name
    return None

def parse_8_column_balance(file_obj_or_path: Any, sheet_name: Optional[str] = None) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Parses any 8-column trial balance from Excel or CSV.
    Returns:
        (df_standardized, metadata_dict)
    """
    metadata = {
        "detected_headers": {},
        "missing_columns": [],
        "total_rows": 0,
        "is_valid": False,
        "error_message": None
    }
    
    # Step 1: Read raw file to find header
    try:
        if hasattr(file_obj_or_path, "name") and file_obj_or_path.name.endswith(".csv"):
            raw_df = pd.read_csv(file_obj_or_path, header=None, dtype=str)
        elif isinstance(file_obj_or_path, str) and file_obj_or_path.endswith(".csv"):
            raw_df = pd.read_csv(file_obj_or_path, header=None, dtype=str)
        else:
            # Excel
            raw_df = pd.read_excel(file_obj_or_path, sheet_name=sheet_name or 0, header=None, dtype=str)
    except Exception as e:
        metadata["error_message"] = f"Error al leer el archivo: {str(e)}"
        return pd.DataFrame(), metadata
    
    # Step 2: Scan rows to identify header row
    header_row_idx = None
    column_mapping = {}
    
    for row_idx in range(min(20, len(raw_df))):
        row_vals = [str(x).strip() for x in raw_df.iloc[row_idx].values if pd.notna(x)]
        matches = {}
        for col_idx, cell_val in enumerate(raw_df.iloc[row_idx]):
            if pd.notna(cell_val):
                matched = match_column(str(cell_val))
                if matched and matched not in matches:
                    matches[matched] = col_idx
        
        # We need at least 4 critical matches (e.g. codigo/nombre, debito/credito or activo/pasivo)
        if len(matches) >= 4 or (("activo" in matches or "deudor" in matches) and len(matches) >= 3):
            header_row_idx = row_idx
            column_mapping = matches
            break
            
    if header_row_idx is None:
        # Fallback: Check if first row works
        for col_idx, cell_val in enumerate(raw_df.iloc[0]):
            matched = match_column(str(cell_val))
            if matched:
                column_mapping[matched] = col_idx
        header_row_idx = 0
    
    metadata["detected_headers"] = column_mapping
    
    # Check minimum columns
    required_cols = ["codigo", "nombre_cuenta"]
    for req in required_cols:
        if req not in column_mapping:
            # If no code column, try to use first column as code if second is name
            if "nombre_cuenta" in column_mapping and "codigo" not in column_mapping:
                for idx in range(len(raw_df.columns)):
                    if idx != column_mapping.get("nombre_cuenta"):
                        column_mapping["codigo"] = idx
                        break
            elif "codigo" in column_mapping and "nombre_cuenta" not in column_mapping:
                for idx in range(len(raw_df.columns)):
                    if idx != column_mapping.get("codigo"):
                        column_mapping["nombre_cuenta"] = idx
                        break
    
    # Extract data slice
    data_df = raw_df.iloc[header_row_idx + 1:].copy()
    
    # Build clean output dataframe
    clean_dict = {}
    
    # Map code & name
    if "codigo" in column_mapping:
        clean_dict["codigo"] = data_df.iloc[:, column_mapping["codigo"]].astype(str).str.strip()
    else:
        clean_dict["codigo"] = [f"CTA_{i+1:04d}" for i in range(len(data_df))]
        
    if "nombre_cuenta" in column_mapping:
        clean_dict["nombre_cuenta"] = data_df.iloc[:, column_mapping["nombre_cuenta"]].astype(str).str.strip()
    else:
        clean_dict["nombre_cuenta"] = clean_dict["codigo"]
    
    # Map 8 columns
    num_cols = ["debito", "credito", "deudor", "acreedor", "activo", "pasivo", "perdida", "ganancia"]
    for col in num_cols:
        if col in column_mapping:
            col_series = data_df.iloc[:, column_mapping[col]]
            clean_dict[col] = col_series.apply(clean_financial_number)
        else:
            clean_dict[col] = 0.0
            metadata["missing_columns"].append(col)
            
    df_clean = pd.DataFrame(clean_dict)
    
    # Step 3: Remove empty rows and Total / Summary rows
    is_not_empty = (df_clean["nombre_cuenta"] != "") & (df_clean["nombre_cuenta"].str.lower() != "nan")
    df_clean = df_clean[is_not_empty].copy()
    
    # Filter summary rows (e.g. "TOTALES", "SUMAS IGUALES", "TOTAL GENERAL")
    summary_pattern = r"(?:total|totales|sumas\s*iguales|resumen|sumas\s*del\s*ejercicio)"
    is_summary = (
        df_clean["nombre_cuenta"].str.lower().str.contains(summary_pattern, regex=True, na=False) |
        df_clean["codigo"].str.lower().str.contains(summary_pattern, regex=True, na=False)
    )
    df_clean = df_clean[~is_summary].copy()
    
    # Step 4: Auto-reconstruct missing columns if partially present
    # If deudor/acreedor are 0 but debito/credito exist:
    if (df_clean["deudor"].sum() == 0 and df_clean["acreedor"].sum() == 0) and (df_clean["debito"].sum() > 0 or df_clean["credito"].sum() > 0):
        diff = df_clean["debito"] - df_clean["credito"]
        df_clean["deudor"] = diff.apply(lambda x: x if x > 0 else 0.0)
        df_clean["acreedor"] = diff.apply(lambda x: -x if x < 0 else 0.0)
    
    metadata["total_rows"] = len(df_clean)
    metadata["is_valid"] = len(df_clean) > 0
    
    return df_clean.reset_index(drop=True), metadata
