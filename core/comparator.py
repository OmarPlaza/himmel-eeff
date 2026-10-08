"""
Multi-Period Comparison & Financial Statement Analysis Engine
Performs Horizontal (Nominal and % variations) and Vertical Analysis across periods.
"""
import pandas as pd
import numpy as np
from typing import Dict, Any, List, Tuple

def compare_classified_balances(df_base: pd.DataFrame, df_current: pd.DataFrame, 
                                threshold_pct: float = 20.0) -> Dict[str, Any]:
    """
    Compares two classified balance DataFrames (Base Year vs Current Year).
    Returns comparative table with nominal and percentage variations and atypical flags.
    """
    # 1. Group by IFRS Category and Rubric
    base_grouped = df_base.groupby(["categoria_ifrs", "rubro_clasificado"])["saldo_final"].sum().reset_index()
    base_grouped.rename(columns={"saldo_final": "saldo_base"}, inplace=True)
    
    curr_grouped = df_current.groupby(["categoria_ifrs", "rubro_clasificado"])["saldo_final"].sum().reset_index()
    curr_grouped.rename(columns={"saldo_final": "saldo_actual"}, inplace=True)
    
    # Merge outer
    merged = pd.merge(base_grouped, curr_grouped, on=["categoria_ifrs", "rubro_clasificado"], how="outer").fillna(0.0)
    
    # Calculate variations
    merged["variacion_nominal"] = merged["saldo_actual"] - merged["saldo_base"]
    
    def calc_pct(row):
        base = row["saldo_base"]
        curr = row["saldo_actual"]
        if base == 0 and curr == 0:
            return 0.0
        if base == 0 and curr > 0:
            return 100.0  # New account
        if base == 0 and curr < 0:
            return -100.0
        return ((curr - base) / abs(base)) * 100.0
        
    merged["variacion_pct"] = merged.apply(calc_pct, axis=1)
    
    # Flag significant / atypical variations
    merged["es_atipica"] = merged["variacion_pct"].abs() > threshold_pct
    
    # 2. Detailed Account Level Comparison
    base_acc = df_base[["codigo", "nombre_cuenta", "rubro_clasificado", "saldo_final"]].copy()
    base_acc.rename(columns={"saldo_final": "saldo_base", "nombre_cuenta": "nombre_base"}, inplace=True)
    
    curr_acc = df_current[["codigo", "nombre_cuenta", "rubro_clasificado", "saldo_final"]].copy()
    curr_acc.rename(columns={"saldo_final": "saldo_actual", "nombre_cuenta": "nombre_actual"}, inplace=True)
    
    merged_acc = pd.merge(base_acc, curr_acc, on="codigo", how="outer").fillna(0.0)
    
    # Fill names
    merged_acc["nombre_cuenta"] = merged_acc.apply(
        lambda r: r["nombre_actual"] if str(r["nombre_actual"]) != "0.0" and str(r["nombre_actual"]) != "" else r["nombre_base"], axis=1
    )
    merged_acc["rubro_clasificado"] = merged_acc.apply(
        lambda r: r["rubro_clasificado_y"] if str(r["rubro_clasificado_y"]) != "0.0" and str(r["rubro_clasificado_y"]) != "" else r["rubro_clasificado_x"], axis=1
    )
    
    merged_acc["variacion_nominal"] = merged_acc["saldo_actual"] - merged_acc["saldo_base"]
    merged_acc["variacion_pct"] = merged_acc.apply(calc_pct, axis=1)
    
    # Account status
    def get_acc_status(r):
        if r["saldo_base"] == 0 and r["saldo_actual"] > 0:
            return "CUENTA_NUEVA"
        if r["saldo_base"] > 0 and r["saldo_actual"] == 0:
            return "CUENTA_DESCONTINUADA"
        if abs(r["variacion_pct"]) > threshold_pct:
            return "VARIACIÓN_SIGNIFICATIVA"
        return "NORMAL"
        
    merged_acc["estado_comparativo"] = merged_acc.apply(get_acc_status, axis=1)
    
    # Clean up columns
    clean_acc_df = merged_acc[[
        "codigo", "nombre_cuenta", "rubro_clasificado", "saldo_base", "saldo_actual", 
        "variacion_nominal", "variacion_pct", "estado_comparativo"
    ]].copy()
    
    # Extract executive highlights
    nuevas = clean_acc_df[clean_acc_df["estado_comparativo"] == "CUENTA_NUEVA"].to_dict(orient="records")
    descontinuadas = clean_acc_df[clean_acc_df["estado_comparativo"] == "CUENTA_DESCONTINUADA"].to_dict(orient="records")
    mayores_alzas = clean_acc_df.sort_values(by="variacion_nominal", ascending=False).head(5).to_dict(orient="records")
    mayores_bajas = clean_acc_df.sort_values(by="variacion_nominal", ascending=True).head(5).to_dict(orient="records")
    
    return {
        "rubros_comparativo_df": merged,
        "cuentas_comparativo_df": clean_acc_df,
        "cuentas_nuevas": nuevas,
        "cuentas_descontinuadas": descontinuadas,
        "top_alzas": mayores_alzas,
        "top_bajas": mayores_bajas
    }
