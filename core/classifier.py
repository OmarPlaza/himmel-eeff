"""
IFRS & NIC 1 Financial Classifier Engine
Robust account code prefix parser and hierarchical mapper for Himmel SpA and standard IFRS.
Uses embedded HIMMEL_MASTER_CATALOG with optional SQLite custom overrides.
Ensures 100% deterministic separation between Assets (1xx), Liabilities (2xx), Equity (3xx), and P&L (4xx/5xx).
"""
import re
import pandas as pd
from typing import Dict, Any, List, Tuple, Optional
from core.himmel_catalog_data import HIMMEL_MASTER_CATALOG

try:
    from core.database import get_connection
except Exception:
    get_connection = None

def classify_by_code_prefix(code: str, name: str, deudor: float, acreedor: float,
                            activo: float, pasivo: float, perdida: float, ganancia: float) -> Tuple[str, str, str, str, str, str]:
    """
    Deterministically determines Section, Category, Rubro, Subrubro, Nota, and Sheet based on account code prefix.
    Returns: (section, category, rubro, subrubro, nota, sheet_target)
    """
    c = str(code).strip()
    norm_name = str(name).strip().lower()
    
    # -------------------------------------------------------------
    # 1. ACTIVOS (Prefix 1xxx)
    # -------------------------------------------------------------
    if c.startswith("1"):
        if c.startswith("101") or c.startswith("102") or c.startswith("103") or c.startswith("110300") or any(w in norm_name for w in ["banco", "caja", "ffmm", "deposito"]):
            return ("ACTIVOS", "ACTIVOS CORRIENTES", "Efectivo y Equivalentes al Efectivo", "Disponible Bancos y Deposito a Plazo", "4.1", "Balance")
        elif c.startswith("106001") or c.startswith("106002") or c.startswith("106005") or "anticipo" in norm_name:
            return ("ACTIVOS", "ACTIVOS CORRIENTES", "Otros Activos No Financieros Corrientes", "Anticipos a Proveedores y Personal", "8", "Balance")
        elif c.startswith("106007") or c.startswith("106009") or (c.startswith("10601") and len(c)>=6) or "relacionad" in norm_name or "himmel" in norm_name:
            return ("ACTIVOS", "ACTIVOS CORRIENTES", "Cuentas por Cobrar a Entidades Relacionadas, Corrientes", "Empresas Relacionadas", "9.1", "Balance")
        elif c.startswith("104") or c.startswith("105") or c.startswith("107") or "cliente" in norm_name or "op kp" in norm_name:
            return ("ACTIVOS", "ACTIVOS CORRIENTES", "Deudores comerciales y Otras Cuentas por Cobrar, Corrientes", "Clientes", "5.1 - 9", "Balance")
        elif c.startswith("110") or c.startswith("111") or "iva credito" in norm_name or "ppm" in norm_name:
            return ("ACTIVOS", "ACTIVOS CORRIENTES", "Activos por Impuestos Corrientes, Corrientes", "Impuestos por Recuperar", "7", "Balance")
        elif c.startswith("1201") or c.startswith("1203") or c.startswith("1204") or c.startswith("1205") or any(w in norm_name for w in ["terreno", "edificio", "maquinaria", "vehiculo"]):
            return ("ACTIVOS", "ACTIVOS NO CORRIENTES", "Propiedades, Planta y Equipo", "Bienes de Uso y Terrenos", "10.1", "Balance")
        elif c.startswith("1202") or "software" in norm_name or "intangible" in norm_name:
            return ("ACTIVOS", "ACTIVOS NO CORRIENTES", "Activos Intangibles Distintos de la Plusvalía", "Software y Licencias", "10.1", "Balance")
        else:
            return ("ACTIVOS", "ACTIVOS CORRIENTES", "Otros Activos No Financieros Corrientes", "Otros Activos", "8", "Balance")

    # -------------------------------------------------------------
    # 2. PASIVOS (Prefix 2xxx)
    # -------------------------------------------------------------
    elif c.startswith("2"):
        if c.startswith("201") or (c.startswith("2104") and "banco" in norm_name) or "prestamo" in norm_name:
            return ("PASIVOS Y PATRIMONIO", "PASIVOS CORRIENTES", "Otros Pasivos Financieros Corrientes", "Créditos Bancarios", "11.1", "Balance")
        elif c.startswith("216") or "aporte" in norm_name:
            return ("PASIVOS Y PATRIMONIO", "PASIVOS CORRIENTES", "Operaciones Pendientes", "Cuentas Aportes Inversionistas", "11.5", "Balance")
        elif c.startswith("2101") or c.startswith("2102") or c.startswith("2103") or any(w in norm_name for w in ["proveedor", "honorario", "iva debito", "retencion"]):
            return ("PASIVOS Y PATRIMONIO", "PASIVOS CORRIENTES", "Pasivos por Impuestos Corrientes, Corrientes", "Proveedores e Impuestos por pagar", "11.2 - 11.3 - 11.4 - 13.1", "Balance")
        elif c.startswith("217") or c.startswith("218") or c.startswith("219") or "provision" in norm_name:
            return ("PASIVOS Y PATRIMONIO", "PASIVOS CORRIENTES", "Otros Pasivos No Financieros Corrientes", "Provisiones Varias", "12.1 - 12.3", "Balance")
        elif c.startswith("220") or c.startswith("221") or "largo plazo" in norm_name or "hipotecario" in norm_name:
            return ("PASIVOS Y PATRIMONIO", "PASIVOS NO CORRIENTES", "Obligaciones Bancarias a Largo Plazo", "Créditos Bancarios", "15.1 - 15.2", "Balance")
        elif c.startswith("222") or "relacionad" in norm_name:
            return ("PASIVOS Y PATRIMONIO", "PASIVOS NO CORRIENTES", "Cuentas Corrientes Empresas Relacionadas", "Empresas Relacionadas por Pagar", "15.3", "Balance")
        else:
            return ("PASIVOS Y PATRIMONIO", "PASIVOS CORRIENTES", "Cuentas Comerciales y Otras Cuentas por Pagar Corrientes", "Otros Pasivos", "11.2", "Balance")

    # -------------------------------------------------------------
    # 3. PATRIMONIO (Prefix 3xxx)
    # -------------------------------------------------------------
    elif c.startswith("3"):
        if c.startswith("3101") or c.startswith("310201") or "capital" in norm_name:
            return ("PASIVOS Y PATRIMONIO", "PATRIMONIO", "Capital Pagado", "Capital Pagado", "16", "Balance")
        elif c.startswith("3103") or any(w in norm_name for w in ["acumulada", "acumulado", "ejercicio", "utilidad"]):
            return ("PASIVOS Y PATRIMONIO", "PATRIMONIO", "Resultados Acumulados", "Utilidades Acumuladas", "16", "Balance")
        else:
            return ("PASIVOS Y PATRIMONIO", "PATRIMONIO", "Otras Reservas", "Reservas", "16", "Balance")

    # -------------------------------------------------------------
    # 4. EGRESOS / GASTOS (Prefix 4xxx) -> EERR
    # -------------------------------------------------------------
    elif c.startswith("4"):
        if c.startswith("401") or c.startswith("402012") or c.startswith("402013") or "costo" in norm_name:
            return ("ESTADO DE RESULTADOS", "EERR", "Costo de Ventas", "Costo de Ventas", "-", "EERR")
        elif c.startswith("405") or c.startswith("406022") or "financier" in norm_name or "interes" in norm_name:
            return ("ESTADO DE RESULTADOS", "EERR", "Costos Financieros", "Costos Financieros", "-", "EERR")
        elif c.startswith("406002") or c.startswith("406003") or "diferencia" in norm_name or "cambio" in norm_name:
            return ("ESTADO DE RESULTADOS", "EERR", "Diferencias de Cambio", "Diferencias de Cambio", "-", "EERR")
        elif c.startswith("405006") or "correccion" in norm_name or "reajuste" in norm_name:
            return ("ESTADO DE RESULTADOS", "EERR", "Corrección Monetaria", "Corrección Monetaria", "-", "EERR")
        else:
            return ("ESTADO DE RESULTADOS", "EERR", "Gastos de Administración y Ventas", "Gastos de Administración", "-", "EERR")

    # -------------------------------------------------------------
    # 5. INGRESOS (Prefix 5xxx) -> EERR
    # -------------------------------------------------------------
    elif c.startswith("5"):
        if c.startswith("502002") or c.startswith("502008") or "diferencia" in norm_name or "cambio" in norm_name:
            return ("ESTADO DE RESULTADOS", "EERR", "Diferencias de Cambio", "Diferencias de Cambio", "-", "EERR")
        else:
            return ("ESTADO DE RESULTADOS", "EERR", "Ingresos de Actividades Ordinarias", "Ingresos de Explotación", "-", "EERR")

    # Fallback based on accounting nature
    if ganancia > 0:
        return ("ESTADO DE RESULTADOS", "EERR", "Ingresos de Actividades Ordinarias", "Ingresos de Explotación", "-", "EERR")
    elif perdida > 0:
        return ("ESTADO DE RESULTADOS", "EERR", "Gastos de Administración y Ventas", "Gastos de Administración", "-", "EERR")
    elif pasivo > 0 or acreedor > 0:
        return ("PASIVOS Y PATRIMONIO", "PASIVOS CORRIENTES", "Cuentas Comerciales y Otras Cuentas por Pagar Corrientes", "Otros Pasivos", "11.2", "Balance")
    else:
        return ("ACTIVOS", "ACTIVOS CORRIENTES", "Otros Activos No Financieros Corrientes", "Otros Activos", "8", "Balance")

def map_and_classify_dataframe(df: pd.DataFrame, company_id: Optional[int] = None) -> pd.DataFrame:
    """
    Enriches 8-column DataFrame with master catalog rules and prefix fallback.
    Always reliable in any cloud or local environment.
    """
    rules = dict(HIMMEL_MASTER_CATALOG)
    
    # Try loading SQLite custom rules if DB is available
    if get_connection is not None:
        try:
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("SELECT account_code, account_name, section, category, rubro, subrubro, nota, sheet_target FROM himmel_catalog")
            for r in cur.fetchall():
                c_code = str(r[0]).strip()
                rules[c_code] = {
                    "name": r[1], "section": r[2], "category": r[3],
                    "rubro": r[4], "subrubro": r[5], "nota": r[6], "sheet": r[7]
                }
            conn.close()
        except Exception:
            pass

    df_result = df.copy()
    
    sections = []
    categorias = []
    rubros = []
    subrubros = []
    notas = []
    statuses = []
    saldos_finales = []
    sheets_target = []
    
    for _, row in df_result.iterrows():
        raw_code = str(row.get("codigo", "")).strip()
        code = raw_code[:-2] if raw_code.endswith(".0") else raw_code
        code = code.lstrip("'").strip()
        name = str(row.get("nombre_cuenta", "")).strip()
        debito = float(row.get("debito", 0))
        credito = float(row.get("credito", 0))
        deudor = float(row.get("deudor", 0))
        acreedor = float(row.get("acreedor", 0))
        activo = float(row.get("activo", 0))
        pasivo = float(row.get("pasivo", 0))
        perdida = float(row.get("perdida", 0))
        ganancia = float(row.get("ganancia", 0))
        
        # Robustly calculate absolute balance
        if deudor != 0:
            saldo_val = abs(deudor)
        elif activo != 0:
            saldo_val = abs(activo)
        elif acreedor != 0:
            saldo_val = abs(acreedor)
        elif pasivo != 0:
            saldo_val = abs(pasivo)
        elif perdida != 0:
            saldo_val = abs(perdida)
        elif ganancia != 0:
            saldo_val = abs(ganancia)
        else:
            saldo_val = abs(debito - credito)
            
        if code in rules:
            cat_info = rules[code]
            sec = cat_info["section"]
            cat = cat_info["category"]
            rub = cat_info["rubro"]
            sub = cat_info["subrubro"]
            nota = cat_info["nota"]
            sheet = cat_info["sheet"]
            status = "IDÉNTICO"
        else:
            sec, cat, rub, sub, nota, sheet = classify_by_code_prefix(
                code, name, deudor, acreedor, activo, pasivo, perdida, ganancia
            )
            status = "HOMOLOGADO"
            
        # Deterministically guarantee sheet target based on accounting hierarchy
        if code.startswith(("4", "5")):
            sheet = "EERR"
            sec = "ESTADO DE RESULTADOS"
        elif code.startswith("1"):
            sheet = "Balance"
            sec = "ACTIVOS"
        elif code.startswith(("2", "3")):
            sheet = "Balance"
            sec = "PASIVOS Y PATRIMONIO"
            
        sections.append(sec)
        categorias.append(cat)
        rubros.append(rub)
        subrubros.append(sub)
        notas.append(nota)
        statuses.append(status)
        saldos_finales.append(saldo_val)
        sheets_target.append(sheet)
        
    df_result["seccion"] = [str(x) if pd.notna(x) else "" for x in sections]
    df_result["categoria_ifrs"] = [str(x) if pd.notna(x) else "" for x in categorias]
    df_result["rubro_clasificado"] = [str(x) if pd.notna(x) else "" for x in rubros]
    df_result["subrubro"] = [str(x) if pd.notna(x) else "" for x in subrubros]
    df_result["nota_auditada"] = [str(x) if pd.notna(x) else "" for x in notas]
    df_result["mapping_status"] = [str(x) if pd.notna(x) else "" for x in statuses]
    df_result["saldo_final"] = saldos_finales
    df_result["sheet_target"] = [str(x) if pd.notna(x) else "" for x in sheets_target]
    df_result["codigo"] = df_result["codigo"].astype(str).fillna("")
    df_result["nombre_cuenta"] = df_result["nombre_cuenta"].astype(str).fillna("")
    
    return df_result

def get_all_rubros_list() -> List[str]:
    rubros = set(v["rubro"] for v in HIMMEL_MASTER_CATALOG.values() if v.get("rubro"))
    return sorted(list(rubros))

def save_custom_mapping(account_code: str, account_name: str, rubro: str, subrubro: str = "", nota: str = ""):
    HIMMEL_MASTER_CATALOG[account_code] = {
        "name": account_name,
        "section": "ACTIVOS" if account_code.startswith("1") else ("PASIVOS Y PATRIMONIO" if account_code.startswith("2") or account_code.startswith("3") else "ESTADO DE RESULTADOS"),
        "category": "ACTIVOS CORRIENTES" if account_code.startswith("1") else ("PASIVOS CORRIENTES" if account_code.startswith("2") else "PATRIMONIO"),
        "rubro": rubro,
        "subrubro": subrubro or rubro,
        "nota": nota,
        "sheet": "Balance" if not (account_code.startswith("4") or account_code.startswith("5")) else "EERR"
    }
    if get_connection is not None:
        try:
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("""
            INSERT OR REPLACE INTO himmel_catalog (account_code, account_name, section, category, rubro, subrubro, nota, sheet_target, line_order)
            VALUES (?, ?, 'PERSONALIZADO', 'ACTIVO_CORRIENTE', ?, ?, ?, 'Balance', 999)
            """, (account_code, account_name, rubro, subrubro, nota))
            conn.commit()
            conn.close()
        except Exception:
            pass

def build_classified_financial_statements(df_classified: pd.DataFrame) -> Dict[str, Any]:
    summary = {
        "activos": {},
        "pasivos_y_patrimonio": {},
        "eerr": {},
        "totals": {
            "total_activo_corriente": 0.0,
            "total_activo_no_corriente": 0.0,
            "total_activos": 0.0,
            "total_pasivo_corriente": 0.0,
            "total_pasivo_no_corriente": 0.0,
            "total_pasivos": 0.0,
            "total_patrimonio": 0.0,
            "total_pasivo_y_patrimonio": 0.0,
            "diferencia_cuadre": 0.0,
            "total_ingresos": 0.0,
            "total_costos": 0.0,
            "margen_bruto": 0.0,
            "total_gastos_admin": 0.0,
            "resultado_operacional": 0.0,
            "total_costos_financieros": 0.0,
            "total_diferencia_cambio": 0.0,
            "total_correccion_monetaria": 0.0,
            "resultado_del_ejercicio": 0.0
        }
    }
    
    for _, row in df_classified.iterrows():
        sec = str(row["seccion"])
        cat = str(row["categoria_ifrs"])
        rub = str(row["rubro_clasificado"])
        sub = str(row["subrubro"])
        nota = str(row["nota_auditada"])
        code = str(row["codigo"])
        name = str(row["nombre_cuenta"])
        val = float(row["saldo_final"])
        sheet = str(row["sheet_target"])
        
        if sheet == "EERR" or "RESULTADOS" in sec or code.startswith("4") or code.startswith("5"):
            target = summary["eerr"]
        elif "ACTIVO" in sec or code.startswith("1"):
            target = summary["activos"]
        else:
            target = summary["pasivos_y_patrimonio"]
            
        if cat not in target:
            target[cat] = {"rubros": {}, "subtotal": 0.0}
        if rub not in target[cat]["rubros"]:
            target[cat]["rubros"][rub] = {"subrubros": {}, "subtotal": 0.0, "nota": nota}
        if sub not in target[cat]["rubros"][rub]["subrubros"]:
            target[cat]["rubros"][rub]["subrubros"][sub] = {"cuentas": [], "subtotal": 0.0, "nota": nota}
            
        target[cat]["rubros"][rub]["subrubros"][sub]["cuentas"].append({
            "codigo": code, "nombre_cuenta": name, "saldo": val, "nota": nota
        })
        target[cat]["rubros"][rub]["subrubros"][sub]["subtotal"] += val
        target[cat]["rubros"][rub]["subtotal"] += val
        target[cat]["subtotal"] += val
        
    act_corr = summary["activos"].get("ACTIVOS CORRIENTES", {}).get("subtotal", 0.0)
    act_no_corr = summary["activos"].get("ACTIVOS NO CORRIENTES", {}).get("subtotal", 0.0)
    tot_act = act_corr + act_no_corr
    
    pas_corr = summary["pasivos_y_patrimonio"].get("PASIVOS CORRIENTES", {}).get("subtotal", 0.0)
    pas_no_corr = summary["pasivos_y_patrimonio"].get("PASIVOS NO CORRIENTES", {}).get("subtotal", 0.0)
    tot_pas = pas_corr + pas_no_corr
    
    tot_pat = summary["pasivos_y_patrimonio"].get("PATRIMONIO", {}).get("subtotal", 0.0)
    tot_pas_pat = tot_pas + tot_pat
    
    # Calculate EERR using flexible matching for rubro names
    is_df = df_classified[(df_classified["sheet_target"] == "EERR") | (df_classified["seccion"] == "ESTADO DE RESULTADOS") | df_classified["codigo"].str.startswith("4") | df_classified["codigo"].str.startswith("5")]
    
    tot_ing = is_df[is_df["rubro_clasificado"].str.contains("Ingreso", case=False, na=False)]["saldo_final"].sum()
    tot_cos = is_df[is_df["rubro_clasificado"].str.contains("Costo de Venta", case=False, na=False)]["saldo_final"].sum()
    margen_bruto = tot_ing - tot_cos
    tot_adm = is_df[is_df["rubro_clasificado"].str.contains("Administra", case=False, na=False)]["saldo_final"].sum()
    res_op = margen_bruto - tot_adm
    tot_fin = is_df[is_df["rubro_clasificado"].str.contains("Financier", case=False, na=False)]["saldo_final"].sum()
    tot_tc = is_df[is_df["rubro_clasificado"].str.contains("Cambio", case=False, na=False)]["saldo_final"].sum()
    tot_cm = is_df[is_df["rubro_clasificado"].str.contains("Monetaria", case=False, na=False)]["saldo_final"].sum()
    res_neto = res_op - tot_fin - tot_tc - tot_cm
    
    summary["totals"].update({
        "total_activo_corriente": act_corr,
        "total_activo_no_corriente": act_no_corr,
        "total_activos": tot_act,
        "total_pasivo_corriente": pas_corr,
        "total_pasivo_no_corriente": pas_no_corr,
        "total_pasivos": tot_pas,
        "total_patrimonio": tot_pat,
        "total_pasivo_y_patrimonio": tot_pas_pat,
        "diferencia_cuadre": tot_act - tot_pas_pat,
        "total_ingresos": tot_ing,
        "total_costos": tot_cos,
        "margen_bruto": margen_bruto,
        "total_gastos_admin": tot_adm,
        "resultado_operacional": res_op,
        "total_costos_financieros": tot_fin,
        "total_diferencia_cambio": tot_tc,
        "total_correccion_monetaria": tot_cm,
        "resultado_del_ejercicio": res_neto
    })
    
    return summary
