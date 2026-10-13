"""
Robust Parser for 8-Column Trial Balances (Balance de 8 Columnas)
Supports Excel (.xlsx, .xls) and CSV formats with automated sheet discovery,
header detection, currency sanitization, and automatic classification.
"""
import io
import os
import re
import pandas as pd
import numpy as np
from typing import Dict, Any, Tuple, Optional, List

# Column aliases mapping (flexible, matching Chilean ERPs like Softland, Defontana, ERP, etc.)
COLUMN_PATTERNS = {
    "codigo": [r"c[oó]d", r"n[o°º]?\s*cta", r"^cuenta$", r"n_cuenta", r"c[oó]digo", r"^id$"],
    "nombre_cuenta": [r"nom", r"descripci[oó]n", r"glosa", r"detalle", r"denominaci[oó]n", r"nombre\s*cuenta"],
    "debito": [r"d[eé]bito", r"debitos", r"cargo", r"cargos"],
    "credito": [r"cr[eé]dito", r"creditos", r"abono", r"abonos"],
    "deudor": [r"deudor", r"saldo\s*d", r"saldos\s*deudor"],
    "acreedor": [r"acreedor", r"saldo\s*a", r"saldos\s*acreedor"],
    "activo": [r"activo", r"activos", r"inventario\s*activo"],
    "pasivo": [r"pasivo", r"pasivos", r"inventario\s*pasivo"],
    "perdida": [r"p[eé]rdida", r"perdidas", r"p[eé]rdidas", r"gasto", r"gastos", r"resultado\s*p[eé]rdida"],
    "ganancia": [r"ganancia", r"ganancias", r"ingreso", r"ingresos", r"resultado\s*ganancia"]
}

def clean_account_code(code_val: Any) -> str:
    """Sanitizes accounting account code, removing trailing .0, spaces, and punctuation."""
    if pd.isna(code_val) or code_val is None:
        return ""
    s = str(code_val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.lstrip("'")
    return s.strip()

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
    if "." in val_str and "," in val_str:
        if val_str.rfind(",") > val_str.rfind("."):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "." in val_str and val_str.count(".") > 1:
        val_str = val_str.replace(".", "")
    elif "," in val_str and val_str.count(",") > 1:
        val_str = val_str.replace(",", "")
    elif "," in val_str and len(val_str.split(",")[-1]) != 2:
        val_str = val_str.replace(",", "")
    elif "," in val_str:
        val_str = val_str.replace(",", ".")
        
    try:
        num = float(val_str)
        return -num if is_negative else num
    except ValueError:
        return 0.0

def match_column(col_name: str) -> Optional[str]:
    """Matches raw column names against standardized financial columns."""
    normalized = str(col_name).strip().lower()
    for canonical_name, patterns in COLUMN_PATTERNS.items():
        for pat in patterns:
            if re.search(pat, normalized):
                return canonical_name
    return None

def parse_8_column_balance(file_obj_or_path: Any, sheet_name: Optional[str] = None) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Parses any 8-column trial balance from Excel or CSV.
    Safely buffers raw bytes to handle Streamlit uploads and local paths identically.
    Scans multiple sheets and finds the header row up to row 50.
    """
    metadata = {
        "detected_headers": {},
        "missing_columns": [],
        "total_rows": 0,
        "is_valid": False,
        "error_message": None,
        "sheet_used": None,
        "header_row_index": None
    }
    
    # Extract raw bytes once to avoid stream exhaustion/closed file errors
    try:
        if hasattr(file_obj_or_path, "read"):
            raw_bytes = file_obj_or_path.read()
            if hasattr(file_obj_or_path, "seek"):
                try:
                    file_obj_or_path.seek(0)
                except Exception:
                    pass
        elif isinstance(file_obj_or_path, (str, os.PathLike)):
            with open(file_obj_or_path, "rb") as f:
                raw_bytes = f.read()
        elif isinstance(file_obj_or_path, bytes):
            raw_bytes = file_obj_or_path
        else:
            raw_bytes = bytes(file_obj_or_path)
    except Exception as e:
        metadata["error_message"] = f"No se pudo leer el archivo: {str(e)}"
        return pd.DataFrame(), metadata

    if not raw_bytes:
        metadata["error_message"] = "El archivo cargado está vacío."
        return pd.DataFrame(), metadata

    is_csv = False
    filename = getattr(file_obj_or_path, "name", str(file_obj_or_path)).lower()
    if filename.endswith(".csv"):
        is_csv = True

    raw_df = None
    chosen_sheet = None
    column_mapping = {}
    header_row_idx = None

    if is_csv:
        for enc in ["utf-8-sig", "latin1", "cp1252", "iso-8859-1"]:
            try:
                candidate_df = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc, header=None, dtype=str)
                for r_idx in range(min(50, len(candidate_df))):
                    matches = {}
                    for c_idx, cell_val in enumerate(candidate_df.iloc[r_idx]):
                        if pd.notna(cell_val):
                            matched = match_column(str(cell_val))
                            if matched and matched not in matches:
                                matches[matched] = c_idx
                    if len(matches) >= 4 or (("activo" in matches or "deudor" in matches) and len(matches) >= 3):
                        header_row_idx = r_idx
                        column_mapping = matches
                        raw_df = candidate_df
                        chosen_sheet = "CSV"
                        break
                if header_row_idx is not None:
                    break
            except Exception:
                continue
    else:
        # Excel parsing
        excel_file = None
        sheets_to_try = [sheet_name] if sheet_name is not None else []
        try:
            excel_file = pd.ExcelFile(io.BytesIO(raw_bytes))
            all_sheets = excel_file.sheet_names
            if sheet_name is not None and sheet_name in all_sheets:
                sheets_to_try = [sheet_name]
            else:
                # Prioritize sheets with 'balance' or '8 col'
                balance_sheets = [s for s in all_sheets if "balance" in str(s).lower() or "8 col" in str(s).lower()]
                sheets_to_try = balance_sheets + [s for s in all_sheets if s not in balance_sheets]
        except Exception as e:
            # Fallback to direct read_excel
            sheets_to_try = [0]

        for s in sheets_to_try:
            try:
                if excel_file is not None:
                    candidate_df = excel_file.parse(sheet_name=s, header=None, dtype=str)
                else:
                    candidate_df = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=s, header=None, dtype=str)
            except Exception:
                continue

            for r_idx in range(min(50, len(candidate_df))):
                matches = {}
                for c_idx, cell_val in enumerate(candidate_df.iloc[r_idx]):
                    if pd.notna(cell_val):
                        matched = match_column(str(cell_val))
                        if matched and matched not in matches:
                            matches[matched] = c_idx

                if len(matches) >= 4 or (("activo" in matches or "deudor" in matches) and len(matches) >= 3):
                    header_row_idx = r_idx
                    column_mapping = matches
                    raw_df = candidate_df
                    chosen_sheet = s
                    break

            if header_row_idx is not None:
                break

    # Fallback to row 0 of first sheet if no headers discovered
    if raw_df is None or header_row_idx is None:
        try:
            if is_csv:
                raw_df = pd.read_csv(io.BytesIO(raw_bytes), header=None, dtype=str)
            else:
                raw_df = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=0, header=None, dtype=str)
            for c_idx, cell_val in enumerate(raw_df.iloc[0]):
                matched = match_column(str(cell_val))
                if matched:
                    column_mapping[matched] = c_idx
            header_row_idx = 0
            chosen_sheet = 0
        except Exception as e:
            metadata["error_message"] = f"No se pudo estructurar el balance: {str(e)}"
            return pd.DataFrame(), metadata

    metadata["detected_headers"] = column_mapping
    metadata["sheet_used"] = chosen_sheet
    metadata["header_row_index"] = header_row_idx

    # Ensure codigo & nombre_cuenta are assigned
    if "codigo" not in column_mapping:
        column_mapping["codigo"] = 0
    if "nombre_cuenta" not in column_mapping:
        column_mapping["nombre_cuenta"] = 1

    # Extract rows below header
    data_df = raw_df.iloc[header_row_idx + 1:].copy()

    clean_dict = {}
    clean_dict["codigo"] = data_df.iloc[:, column_mapping["codigo"]].apply(clean_account_code)
    clean_dict["nombre_cuenta"] = data_df.iloc[:, column_mapping["nombre_cuenta"]].astype(str).str.strip()

    num_cols = ["debito", "credito", "deudor", "acreedor", "activo", "pasivo", "perdida", "ganancia"]
    for col in num_cols:
        if col in column_mapping and column_mapping[col] < len(data_df.columns):
            clean_dict[col] = data_df.iloc[:, column_mapping[col]].apply(clean_financial_number)
        else:
            clean_dict[col] = 0.0
            metadata["missing_columns"].append(col)

    df_clean = pd.DataFrame(clean_dict)

    # Filter out empty or header artifact rows
    is_not_empty = (df_clean["nombre_cuenta"] != "") & (df_clean["nombre_cuenta"].str.lower() != "nan")
    df_clean = df_clean[is_not_empty].copy()

    # Filter summary rows (e.g. "TOTALES", "SUMAS IGUALES", "TOTAL GENERAL")
    summary_pattern = r"(?:total|totales|sumas\s*iguales|resumen|sumas\s*del\s*ejercicio)"
    is_summary = (
        df_clean["nombre_cuenta"].str.lower().str.contains(summary_pattern, regex=True, na=False) |
        df_clean["codigo"].str.lower().str.contains(summary_pattern, regex=True, na=False)
    )
    df_clean = df_clean[~is_summary].copy()

    # Filter rows with no code and no numbers
    has_values = (df_clean["debito"] != 0) | (df_clean["credito"] != 0) | (df_clean["deudor"] != 0) | (df_clean["acreedor"] != 0) | (df_clean["activo"] != 0) | (df_clean["pasivo"] != 0) | (df_clean["perdida"] != 0) | (df_clean["ganancia"] != 0)
    has_code = df_clean["codigo"] != ""
    df_clean = df_clean[has_values | has_code].copy()

    # Auto-reconstruct deudor/acreedor if 0 but debito/credito exist
    if (df_clean["deudor"].sum() == 0 and df_clean["acreedor"].sum() == 0) and (df_clean["debito"].sum() > 0 or df_clean["credito"].sum() > 0):
        diff = df_clean["debito"] - df_clean["credito"]
        df_clean["deudor"] = diff.apply(lambda x: x if x > 0 else 0.0)
        df_clean["acreedor"] = diff.apply(lambda x: -x if x < 0 else 0.0)

    metadata["total_rows"] = len(df_clean)
    metadata["is_valid"] = len(df_clean) > 0

    return df_clean.reset_index(drop=True), metadata
