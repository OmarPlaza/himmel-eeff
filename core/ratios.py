"""
CFO Financial Ratios & Key Performance Indicators (KPIs) Engine
Calculates Liquidity, Solvency, Profitability and Operational DuPont metrics.
"""
from typing import Dict, Any

def calculate_financial_ratios(statements: Dict[str, Any]) -> Dict[str, Any]:
    """
    Computes key financial ratios from consolidated IFRS statements.
    """
    t = statements.get("totals", {})
    bs = statements.get("balance_sheet", {})
    
    act_corr = t.get("total_activo_corriente", 0.0)
    act_no_corr = t.get("total_activo_no_corriente", 0.0)
    tot_act = t.get("total_activos", 0.0)
    
    pas_corr = t.get("total_pasivo_corriente", 0.0)
    pas_no_corr = t.get("total_pasivo_no_corriente", 0.0)
    tot_pas = t.get("total_pasivos", 0.0)
    
    tot_pat = t.get("total_patrimonio", 0.0)
    
    ingresos = t.get("total_ingresos", 0.0)
    margen_bruto = t.get("margen_bruto", 0.0)
    res_op = t.get("resultado_operacional", 0.0)
    res_neto = t.get("resultado_del_ejercicio", 0.0)
    
    # Extract inventory from Activo Corriente
    act_corr_rubros = bs.get("ACTIVO_CORRIENTE", {}).get("rubros", {})
    inventarios = act_corr_rubros.get("Inventarios", 0.0)
    disponible = act_corr_rubros.get("Efectivo y Equivalentes de Efectivo", 0.0)
    
    # 1. Ratios de Liquidez
    razon_corriente = (act_corr / pas_corr) if pas_corr > 0 else 0.0
    prueba_acida = ((act_corr - inventarios) / pas_corr) if pas_corr > 0 else 0.0
    ratio_efectivo = (disponible / pas_corr) if pas_corr > 0 else 0.0
    capital_de_trabajo = act_corr - pas_corr
    
    # 2. Ratios de Endeudamiento y Solvencia
    leverage_financiero = (tot_pas / tot_pat) if tot_pat > 0 else 0.0
    endeudamiento_activo = (tot_pas / tot_act) if tot_act > 0 else 0.0
    calidad_deuda = (pas_corr / tot_pas) if tot_pas > 0 else 0.0
    solvencia_patrimonial = (tot_pat / tot_act) if tot_act > 0 else 0.0
    
    # 3. Ratios de Rentabilidad
    roe = (res_neto / tot_pat * 100.0) if tot_pat > 0 else 0.0
    roa = (res_neto / tot_act * 100.0) if tot_act > 0 else 0.0
    margen_bruto_pct = (margen_bruto / ingresos * 100.0) if ingresos > 0 else 0.0
    margen_operacional_pct = (res_op / ingresos * 100.0) if ingresos > 0 else 0.0
    margen_neto_pct = (res_neto / ingresos * 100.0) if ingresos > 0 else 0.0
    
    return {
        "liquidez": {
            "razon_corriente": {
                "valor": round(razon_corriente, 2),
                "unidad": "veces",
                "interpretacion": "Capacidad de cubrir deudas a corto plazo con activos corrientes.",
                "estado": "SALUDABLE" if razon_corriente >= 1.5 else ("PRECAUCIÓN" if razon_corriente >= 1.0 else "CRÍTICO")
            },
            "prueba_acida": {
                "valor": round(prueba_acida, 2),
                "unidad": "veces",
                "interpretacion": "Capacidad de cubrir pasivos de corto plazo sin depender de la venta de inventario.",
                "estado": "SALUDABLE" if prueba_acida >= 1.0 else ("PRECAUCIÓN" if prueba_acida >= 0.8 else "CRÍTICO")
            },
            "capital_trabajo": {
                "valor": round(capital_de_trabajo, 2),
                "unidad": "$",
                "interpretacion": "Excedente de fondos líquidos de operación (Activo Corriente - Pasivo Corriente).",
                "estado": "SALUDABLE" if capital_de_trabajo > 0 else "CRÍTICO"
            },
            "ratio_efectivo": {
                "valor": round(ratio_efectivo, 2),
                "unidad": "veces",
                "interpretacion": "Disponibilidad de caja inmediata frente a obligaciones corrientes."
            }
        },
        "solvencia": {
            "leverage_total": {
                "valor": round(leverage_financiero, 2),
                "unidad": "veces",
                "interpretacion": "Proporción de deuda con terceros frente al patrimonio de los accionistas.",
                "estado": "SALUDABLE" if leverage_financiero <= 1.5 else ("PRECAUCIÓN" if leverage_financiero <= 2.5 else "ELEVADO")
            },
            "endeudamiento_activo": {
                "valor": round(endeudamiento_activo * 100.0, 1),
                "unidad": "%",
                "interpretacion": "Porcentaje de los activos totales financiado con deuda."
            },
            "calidad_deuda_cp": {
                "valor": round(calidad_deuda * 100.0, 1),
                "unidad": "%",
                "interpretacion": "Porcentaje de la deuda total que vence en el corto plazo (< 1 año)."
            },
            "autonomia_financiera": {
                "valor": round(solvencia_patrimonial * 100.0, 1),
                "unidad": "%",
                "interpretacion": "Porcentaje del activo financiado con recursos propios."
            }
        },
        "rentabilidad": {
            "roe": {
                "valor": round(roe, 2),
                "unidad": "%",
                "interpretacion": "Retorno sobre el Patrimonio (Return on Equity).",
                "estado": "POSITIVO" if roe > 10.0 else ("MODERADO" if roe > 0 else "NEGATIVO")
            },
            "roa": {
                "valor": round(roa, 2),
                "unidad": "%",
                "interpretacion": "Retorno sobre los Activos (Return on Assets).",
                "estado": "POSITIVO" if roa > 5.0 else ("MODERADO" if roa > 0 else "NEGATIVO")
            },
            "margen_bruto": {
                "valor": round(margen_bruto_pct, 1),
                "unidad": "%",
                "interpretacion": "Margen de ganancia bruta sobre las ventas."
            },
            "margen_operacional": {
                "valor": round(margen_operacional_pct, 1),
                "unidad": "%",
                "interpretacion": "Margen de utilidad de las operaciones principales."
            },
            "margen_neto": {
                "valor": round(margen_neto_pct, 1),
                "unidad": "%",
                "interpretacion": "Rendimiento final neto por cada peso vendido."
            }
        }
    }
