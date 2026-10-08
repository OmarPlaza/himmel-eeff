"""
Deterministic Mathematical & Quality Validator for 8-Column Balances
Ensures IFRS rigor, zero-tolerance discrepancy auditing, and inverted balance detection.
"""
import pandas as pd
import numpy as np
from typing import Dict, Any, List, Tuple

def validate_8_column_balance(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Executes deep accounting validations on the 8-column trial balance.
    Returns audit metrics, balances, differences, and anomaly alerts.
    """
    results = {
        "is_mathematically_sound": True,
        "status_code": "OK",  # OK, WARNING, CRITICAL_ERROR
        "sums": {},
        "differences": {},
        "alerts": [],
        "reclassifications_suggested": [],
        "summary_kpis": {}
    }
    
    if df.empty:
        results["is_mathematically_sound"] = False
        results["status_code"] = "CRITICAL_ERROR"
        results["alerts"].append("El balance está vacío o no contiene filas contables válidas.")
        return results
        
    # Calculate column totals
    sum_debito = float(df["debito"].sum())
    sum_credito = float(df["credito"].sum())
    sum_deudor = float(df["deudor"].sum())
    sum_acreedor = float(df["acreedor"].sum())
    sum_activo = float(df["activo"].sum())
    sum_pasivo = float(df["pasivo"].sum())
    sum_perdida = float(df["perdida"].sum())
    sum_ganancia = float(df["ganancia"].sum())
    
    results["sums"] = {
        "debito": sum_debito,
        "credito": sum_credito,
        "deudor": sum_deudor,
        "acreedor": sum_acreedor,
        "activo": sum_activo,
        "pasivo": sum_pasivo,
        "perdida": sum_perdida,
        "ganancia": sum_ganancia
    }
    
    # Tolerances
    EPSILON = 0.01  # Tolerance for rounding cents
    
    # 1. Sumas Débito vs Crédito
    diff_sumas = abs(sum_debito - sum_credito)
    results["differences"]["sumas_deb_cred"] = round(diff_sumas, 2)
    if diff_sumas > EPSILON and (sum_debito > 0 or sum_credito > 0):
        results["is_mathematically_sound"] = False
        results["status_code"] = "CRITICAL_ERROR"
        results["alerts"].append(f"Descuadre en Sumas: Total Débitos ({sum_debito:,.2f}) != Total Créditos ({sum_credito:,.2f}). Diferencia = {diff_sumas:,.2f}")
        
    # 2. Saldos Deudor vs Acreedor
    diff_saldos = abs(sum_deudor - sum_acreedor)
    results["differences"]["saldos_deud_acred"] = round(diff_saldos, 2)
    if diff_saldos > EPSILON and (sum_deudor > 0 or sum_acreedor > 0):
        results["is_mathematically_sound"] = False
        results["status_code"] = "CRITICAL_ERROR"
        results["alerts"].append(f"Descuadre en Saldos: Total Deudor ({sum_deudor:,.2f}) != Total Acreedor ({sum_acreedor:,.2f}). Diferencia = {diff_saldos:,.2f}")
        
    # 3. Resultado del Ejercicio (Inventario vs Resultados)
    res_inventario = sum_activo - sum_pasivo
    res_resultados = sum_ganancia - sum_perdida
    diff_resultado = abs(res_inventario - res_resultados)
    results["differences"]["resultado_ejercicio_diff"] = round(diff_resultado, 2)
    
    # Check if Activo/Pasivo/Perdida/Ganancia are used
    has_classification = (sum_activo > 0 or sum_pasivo > 0 or sum_perdida > 0 or sum_ganancia > 0)
    if has_classification:
        if diff_resultado > EPSILON:
            results["is_mathematically_sound"] = False
            results["status_code"] = "CRITICAL_ERROR"
            results["alerts"].append(
                f"Descuadre de Resultado: (Activo - Pasivo = {res_inventario:,.2f}) != (Ganancia - Pérdida = {res_resultados:,.2f}). Diferencia = {diff_resultado:,.2f}"
            )
        else:
            results["summary_kpis"]["resultado_ejercicio"] = res_resultados
            results["summary_kpis"]["tipo_resultado"] = "UTILIDAD" if res_resultados >= 0 else "PÉRDIDA"
    else:
        results["alerts"].append("El balance no tiene columnas de Inventario y Resultados clasificadas. Se requerirá clasificación completa basada en saldos.")
        
    # 4. Check for Inverted Balances & Anomaly Detection
    for _, row in df.iterrows():
        code = str(row.get("codigo", ""))
        name = str(row.get("nombre_cuenta", "")).lower()
        deudor = float(row.get("deudor", 0))
        acreedor = float(row.get("acreedor", 0))
        activo = float(row.get("activo", 0))
        pasivo = float(row.get("pasivo", 0))
        
        # Bank account with credit balance -> Overdraft (Sobregiro)
        if ("banco" in name or "cuenta corriente" in name) and (acreedor > 0 or pasivo > 0) and deudor == 0:
            val = acreedor if acreedor > 0 else pasivo
            results["reclassifications_suggested"].append({
                "codigo": code,
                "nombre_cuenta": row.get("nombre_cuenta", ""),
                "saldo": val,
                "diagnostico": "Cuenta bancaria con saldo acreedor",
                "propuesta": "Reclasificar a Sobregiros Bancarios (Otros Pasivos Financieros Corrientes)",
                "severidad": "MEDIA"
            })
            
        # Trade Receivables / Clientes with credit balance -> Customer Advance (Anticipo Clientes)
        if ("cliente" in name or "deudores por venta" in name) and (acreedor > 0 or pasivo > 0) and deudor == 0:
            val = acreedor if acreedor > 0 else pasivo
            results["reclassifications_suggested"].append({
                "codigo": code,
                "nombre_cuenta": row.get("nombre_cuenta", ""),
                "saldo": val,
                "diagnostico": "Cuenta de Clientes con saldo acreedor",
                "propuesta": "Reclasificar a Anticipos de Clientes (Otros Pasivos No Financieros Corrientes)",
                "severidad": "MEDIA"
            })
            
        # Trade Payables / Proveedores with debit balance -> Supplier Advance (Anticipo Proveedores)
        if ("proveedor" in name or "acreedores" in name) and (deudor > 0 or activo > 0) and acreedor == 0:
            val = deudor if deudor > 0 else activo
            results["reclassifications_suggested"].append({
                "codigo": code,
                "nombre_cuenta": row.get("nombre_cuenta", ""),
                "saldo": val,
                "diagnostico": "Cuenta de Proveedores con saldo deudor",
                "propuesta": "Reclasificar a Anticipos a Proveedores (Otros Activos No Financieros Corrientes)",
                "severidad": "MEDIA"
            })
            
        # Negative balances in columns where positive is expected
        if (deudor < 0 or acreedor < 0 or activo < 0 or pasivo < 0):
            results["alerts"].append(f"Cuenta '{row.get('nombre_cuenta')}' tiene valores negativos explícitos en sus columnas.")
            if results["status_code"] == "OK":
                results["status_code"] = "WARNING"
                
    if results["reclassifications_suggested"] and results["status_code"] == "OK":
        results["status_code"] = "WARNING"
        
    return results
