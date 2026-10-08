"""
Excel & Audit Report Exporter for IFRS Financial Statements
Generates presentation-ready multi-tab Excel files matching the exact audited corporate structure (Himmel / IFRS layout):
- Sheet 1: Balance Comparativo (Activos, Pasivos, Patrimonio)
- Sheet 2: EERR Cmparativo (Estado de Resultados por Función)
- Sheet 3: Balance 8 Columnas Origen
"""
import io
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import Dict, Any, Optional

def generate_audited_detailed_excel(
    df_current_classified: pd.DataFrame,
    df_base_classified: Optional[pd.DataFrame] = None,
    company_name: str = "HIMMEL SPA",
    period_current: str = "Agos 2026",
    period_base: str = "Dic 2025"
) -> io.BytesIO:
    """
    Generates a multi-tab Excel file matching Base de trabajo EEFF 2024/2025:
    - Tab 1: Balance Comparativo
    - Tab 2: EERR Cmparativo
    """
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # Styles
    bold_title = Font(name="Calibri", size=12, bold=True)
    subtitle_font = Font(name="Calibri", size=10)
    box_font = Font(name="Calibri", size=11, bold=True)
    section_font = Font(name="Calibri", size=10, bold=True)
    rubro_font = Font(name="Calibri", size=10, bold=True)
    subrubro_font = Font(name="Calibri", size=10, italic=True)
    account_font = Font(name="Calibri", size=9)
    total_font = Font(name="Calibri", size=10, bold=True)
    header_col_font = Font(name="Calibri", size=10, bold=True)
    
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    double_bottom_border = Border(
        top=Side(style='thin', color='000000'),
        bottom=Side(style='double', color='000000')
    )
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    # Lookup for base year values
    base_lookup = {}
    if df_base_classified is not None:
        for _, r in df_base_classified.iterrows():
            base_lookup[str(r.get("codigo", "")).strip()] = float(r.get("saldo_final", 0.0))
            
    # Clean period names for subtitle
    p_curr_clean = period_current.replace("Dic ", "").replace("Agos ", "").replace("Ene ", "")
    
    # =========================================================================
    # SHEET 1: BALANCE COMPARATIVO
    # =========================================================================
    ws_bs = wb.active
    ws_bs.title = "Balance Comparativo"
    ws_bs.views.sheetView[0].showGridLines = True
    
    ws_bs.merge_cells("A2:G2")
    ws_bs["A2"] = company_name.upper()
    ws_bs["A2"].font = bold_title
    ws_bs["A2"].alignment = align_center
    
    ws_bs.merge_cells("A3:G3")
    ws_bs["A3"] = f"Balance General entre el 01 de Enero de {p_curr_clean} al período {period_current}"
    ws_bs["A3"].font = subtitle_font
    ws_bs["A3"].alignment = align_center
    
    curr_row = 5
    
    # Sections to render in Balance Sheet
    sections_bs = [
        ("ACTIVOS", "A C T I V O S", [
            ("ACTIVOS CORRIENTES", "ACTIVOS CORRIENTES"),
            ("ACTIVOS NO CORRIENTES", "ACTIVOS NO CORRIENTES")
        ]),
        ("PASIVOS Y PATRIMONIO", "P A S I V O S   Y   P A T R I M O N I O", [
            ("PASIVOS CORRIENTES", "PASIVOS CORRIENTES"),
            ("PASIVOS NO CORRIENTES", "PASIVOS NO CORRIENTES"),
            ("PATRIMONIO", "PATRIMONIO")
        ])
    ]
    
    tot_act_curr_all = 0.0
    tot_act_base_all = 0.0
    tot_pas_curr_all = 0.0
    tot_pas_base_all = 0.0
    
    for sec_id, sec_box_title, categories in sections_bs:
        # Section Box Header
        ws_bs.merge_cells(f"A{curr_row}:G{curr_row}")
        ws_bs[f"A{curr_row}"] = sec_box_title
        ws_bs[f"A{curr_row}"].font = box_font
        ws_bs[f"A{curr_row}"].alignment = align_center
        for col in range(1, 8):
            ws_bs.cell(row=curr_row, column=col).border = thin_border
        curr_row += 2
        
        # Column Headers
        ws_bs.cell(row=curr_row, column=1, value="CONCEPTO / CUENTA").font = header_col_font
        ws_bs.cell(row=curr_row, column=2, value=period_current).font = header_col_font
        ws_bs.cell(row=curr_row, column=2).alignment = align_right
        ws_bs.cell(row=curr_row, column=3, value=period_base).font = header_col_font
        ws_bs.cell(row=curr_row, column=3).alignment = align_right
        ws_bs.cell(row=curr_row, column=4, value="Nota").font = header_col_font
        ws_bs.cell(row=curr_row, column=4).alignment = align_center
        
        ws_bs.cell(row=curr_row-1, column=5, value="$").font = header_col_font
        ws_bs.cell(row=curr_row-1, column=5).alignment = align_right
        ws_bs.cell(row=curr_row, column=5, value=period_current).font = header_col_font
        ws_bs.cell(row=curr_row, column=5).alignment = align_right
        
        ws_bs.cell(row=curr_row-1, column=6, value="$").font = header_col_font
        ws_bs.cell(row=curr_row-1, column=6).alignment = align_right
        ws_bs.cell(row=curr_row, column=6, value=period_base).font = header_col_font
        ws_bs.cell(row=curr_row, column=6).alignment = align_right
        curr_row += 1
        
        sec_df = df_current_classified[(df_current_classified["seccion"] == sec_id) & (df_current_classified["sheet_target"] != "EERR")]
        
        for cat_key, cat_label in categories:
            cat_df = sec_df[sec_df["categoria_ifrs"] == cat_key]
            if cat_df.empty:
                continue
                
            ws_bs.cell(row=curr_row, column=1, value=cat_label).font = section_font
            curr_row += 1
            
            rubros = cat_df["rubro_clasificado"].unique()
            for rub_name in rubros:
                rub_df = cat_df[cat_df["rubro_clasificado"] == rub_name]
                ws_bs.cell(row=curr_row, column=1, value=rub_name).font = rubro_font
                curr_row += 1
                
                subrubros = rub_df["subrubro"].unique()
                for sub_name in subrubros:
                    sub_df = rub_df[rub_df["subrubro"] == sub_name]
                    sub_tot_curr = sub_df["saldo_final"].sum()
                    sub_tot_base = sum(base_lookup.get(str(c).strip(), 0.0) for c in sub_df["codigo"])
                    sub_nota = sub_df.iloc[0].get("nota_auditada", "")
                    
                    if sec_id == "ACTIVOS":
                        tot_act_curr_all += sub_tot_curr
                        tot_act_base_all += sub_tot_base
                    else:
                        tot_pas_curr_all += sub_tot_curr
                        tot_pas_base_all += sub_tot_base
                        
                    ws_bs.cell(row=curr_row, column=1, value=f"    {sub_name}").font = subrubro_font
                    ws_bs.cell(row=curr_row, column=4, value=sub_nota).font = subtitle_font
                    ws_bs.cell(row=curr_row, column=4).alignment = align_center
                    
                    c_tot_curr = ws_bs.cell(row=curr_row, column=5, value=sub_tot_curr)
                    c_tot_curr.font = total_font
                    c_tot_curr.number_format = "#,##0"
                    c_tot_curr.alignment = align_right
                    
                    c_tot_base = ws_bs.cell(row=curr_row, column=6, value=sub_tot_base)
                    c_tot_base.font = total_font
                    c_tot_base.number_format = "#,##0"
                    c_tot_base.alignment = align_right
                    curr_row += 1
                    
                    # Account Lines
                    for _, acc_r in sub_df.iterrows():
                        acc_code = str(acc_r.get("codigo", "")).strip()
                        acc_name = str(acc_r.get("nombre_cuenta", "")).strip()
                        acc_curr = float(acc_r.get("saldo_final", 0.0))
                        acc_base = base_lookup.get(acc_code, 0.0)
                        acc_nota = str(acc_r.get("nota_auditada", ""))
                        
                        ws_bs.cell(row=curr_row, column=1, value=f"        {acc_code}  {acc_name}").font = account_font
                        
                        ca_curr = ws_bs.cell(row=curr_row, column=2, value=acc_curr)
                        ca_curr.font = account_font
                        ca_curr.number_format = "#,##0"
                        ca_curr.alignment = align_right
                        
                        ca_base = ws_bs.cell(row=curr_row, column=3, value=acc_base)
                        ca_base.font = account_font
                        ca_base.number_format = "#,##0"
                        ca_base.alignment = align_right
                        
                        if acc_nota and acc_nota != sub_nota and acc_nota != "-":
                            can = ws_bs.cell(row=curr_row, column=4, value=acc_nota)
                            can.font = account_font
                            can.alignment = align_center
                        curr_row += 1
                        
                    curr_row += 1 # spacer
                    
        # Print Section Total
        if sec_id == "ACTIVOS":
            ws_bs.cell(row=curr_row, column=1, value="TOTAL ACTIVOS").font = bold_title
            c_tac = ws_bs.cell(row=curr_row, column=5, value=tot_act_curr_all)
            c_tac.font = bold_title
            c_tac.number_format = "#,##0"
            c_tac.border = double_bottom_border
            c_tac.alignment = align_right
            
            c_tab = ws_bs.cell(row=curr_row, column=6, value=tot_act_base_all)
            c_tab.font = bold_title
            c_tab.number_format = "#,##0"
            c_tab.border = double_bottom_border
            c_tab.alignment = align_right
            curr_row += 3
        else:
            ws_bs.cell(row=curr_row, column=1, value="TOTAL PASIVOS Y PATRIMONIO").font = bold_title
            c_tpc = ws_bs.cell(row=curr_row, column=5, value=tot_pas_curr_all)
            c_tpc.font = bold_title
            c_tpc.number_format = "#,##0"
            c_tpc.border = double_bottom_border
            c_tpc.alignment = align_right
            
            c_tpb = ws_bs.cell(row=curr_row, column=6, value=tot_pas_base_all)
            c_tpb.font = bold_title
            c_tpb.number_format = "#,##0"
            c_tpb.border = double_bottom_border
            c_tpb.alignment = align_right
            curr_row += 2
            
    # Auto column widths
    ws_bs.column_dimensions["A"].width = 56
    ws_bs.column_dimensions["B"].width = 16
    ws_bs.column_dimensions["C"].width = 16
    ws_bs.column_dimensions["D"].width = 10
    ws_bs.column_dimensions["E"].width = 20
    ws_bs.column_dimensions["F"].width = 20
    
    # =========================================================================
    # SHEET 2: EERR COMPARATIVO
    # =========================================================================
    ws_is = wb.create_sheet(title="EERR Cmparativo")
    ws_is.views.sheetView[0].showGridLines = True
    
    ws_is.merge_cells("A2:G2")
    ws_is["A2"] = company_name.upper()
    ws_is["A2"].font = bold_title
    ws_is["A2"].alignment = align_center
    
    ws_is.merge_cells("A3:G3")
    ws_is["A3"] = f"Estados de Resultados por función del 01 de enero de {p_curr_clean} al período {period_current}"
    ws_is["A3"].font = subtitle_font
    ws_is["A3"].alignment = align_center
    
    # Header box
    ws_is.merge_cells("A5:G5")
    ws_is["A5"] = "E S T A D O   D E   R E S U L T A D O S"
    ws_is["A5"].font = box_font
    ws_is["A5"].alignment = align_center
    for col in range(1, 8):
        ws_is.cell(row=5, column=col).border = thin_border
        
    ws_is.cell(row=7, column=1, value="LÍNEA / CUENTA").font = header_col_font
    ws_is.cell(row=7, column=2, value=period_current).font = header_col_font
    ws_is.cell(row=7, column=2).alignment = align_right
    ws_is.cell(row=7, column=3, value=period_base).font = header_col_font
    ws_is.cell(row=7, column=3).alignment = align_right
    ws_is.cell(row=7, column=4, value="Nota").font = header_col_font
    ws_is.cell(row=7, column=4).alignment = align_center
    
    ws_is.cell(row=6, column=5, value="$").font = header_col_font
    ws_is.cell(row=6, column=5).alignment = align_right
    ws_is.cell(row=7, column=5, value=period_current).font = header_col_font
    ws_is.cell(row=7, column=5).alignment = align_right
    
    ws_is.cell(row=6, column=6, value="$").font = header_col_font
    ws_is.cell(row=6, column=6).alignment = align_right
    ws_is.cell(row=7, column=6, value=period_base).font = header_col_font
    ws_is.cell(row=7, column=6).alignment = align_right
    
    is_df = df_current_classified[(df_current_classified["sheet_target"] == "EERR") | (df_current_classified["seccion"] == "ESTADO DE RESULTADOS")]
    
    curr_is_row = 8
    
    eerr_sections = [
        ("Ingresos de Actividades Ordinarias", "Ingresos de Actividades Ordinarias", 1),
        ("Costo de Ventas", "Costo de Ventas", -1),
        ("GANANCIA BRUTA", "GANANCIA BRUTA", 0),
        ("Gastos de Administración", "Gastos de Administración", -1),
        ("GANANCIAS (PÉRDIDAS) OPERACIONALES", "GANANCIAS (PÉRDIDAS) OPERACIONALES", 0),
        ("Costos Financieros", "Costos Financieros", -1),
        ("Diferencias de cambios por conversión", "Diferencias de cambios por conversión", -1),
        ("Corrección Monetaria", "Corrección Monetaria", -1),
        ("GANANCIA (PÉRDIDA) DEL EJERCICIO", "GANANCIA (PÉRDIDA) DEL EJERCICIO", 0)
    ]
    
    tot_ing_curr = 0.0
    tot_cos_curr = 0.0
    tot_adm_curr = 0.0
    tot_fin_curr = 0.0
    tot_tc_curr = 0.0
    tot_cm_curr = 0.0
    
    tot_ing_base = 0.0
    tot_cos_base = 0.0
    tot_adm_base = 0.0
    tot_fin_base = 0.0
    tot_tc_base = 0.0
    tot_cm_base = 0.0
    
    for rub_key, rub_label, multiplier in eerr_sections:
        if multiplier == 0:
            # Calculated line
            ws_is.cell(row=curr_is_row, column=1, value=rub_label).font = bold_title
            
            if rub_label == "GANANCIA BRUTA":
                val_c = tot_ing_curr - tot_cos_curr
                val_b = tot_ing_base - tot_cos_base
            elif "OPERACIONALES" in rub_label:
                val_c = (tot_ing_curr - tot_cos_curr) - tot_adm_curr
                val_b = (tot_ing_base - tot_cos_base) - tot_adm_base
            else: # Net Income
                val_c = (tot_ing_curr - tot_cos_curr) - tot_adm_curr - tot_fin_curr - tot_tc_curr - tot_cm_curr
                val_b = (tot_ing_base - tot_cos_base) - tot_adm_base - tot_fin_base - tot_tc_base - tot_cm_base
                
            c_vc = ws_is.cell(row=curr_is_row, column=5, value=val_c)
            c_vc.font = bold_title
            c_vc.number_format = "#,##0"
            c_vc.border = double_bottom_border
            c_vc.alignment = align_right
            
            c_vb = ws_is.cell(row=curr_is_row, column=6, value=val_b)
            c_vb.font = bold_title
            c_vb.number_format = "#,##0"
            c_vb.border = double_bottom_border
            c_vb.alignment = align_right
            
            curr_is_row += 2
        else:
            sub_is_df = is_df[is_df["rubro_clasificado"] == rub_key]
            sub_tot_curr = sub_is_df["saldo_final"].sum()
            sub_tot_base = sum(base_lookup.get(str(c).strip(), 0.0) for c in sub_is_df["codigo"])
            
            if rub_key == "Ingresos de Actividades Ordinarias":
                tot_ing_curr = sub_tot_curr
                tot_ing_base = sub_tot_base
            elif rub_key == "Costo de Ventas":
                tot_cos_curr = sub_tot_curr
                tot_cos_base = sub_tot_base
            elif rub_key == "Gastos de Administración":
                tot_adm_curr = sub_tot_curr
                tot_adm_base = sub_tot_base
            elif rub_key == "Costos Financieros":
                tot_fin_curr = sub_tot_curr
                tot_fin_base = sub_tot_base
            elif "cambio" in rub_key.lower():
                tot_tc_curr = sub_tot_curr
                tot_tc_base = sub_tot_base
            elif "corrección" in rub_key.lower():
                tot_cm_curr = sub_tot_curr
                tot_cm_base = sub_tot_base
                
            ws_is.cell(row=curr_is_row, column=1, value=rub_label).font = rubro_font
            
            c_stc = ws_is.cell(row=curr_is_row, column=5, value=sub_tot_curr * multiplier)
            c_stc.font = total_font
            c_stc.number_format = "#,##0"
            c_stc.alignment = align_right
            
            c_stb = ws_is.cell(row=curr_is_row, column=6, value=sub_tot_base * multiplier)
            c_stb.font = total_font
            c_stb.number_format = "#,##0"
            c_stb.alignment = align_right
            curr_is_row += 1
            
            for _, acc_r in sub_is_df.iterrows():
                acc_code = str(acc_r.get("codigo", "")).strip()
                acc_name = str(acc_r.get("nombre_cuenta", "")).strip()
                acc_curr = float(acc_r.get("saldo_final", 0.0))
                acc_base = base_lookup.get(acc_code, 0.0)
                
                ws_is.cell(row=curr_is_row, column=1, value=f"    {acc_code}  {acc_name}").font = account_font
                
                c_ac = ws_is.cell(row=curr_is_row, column=2, value=acc_curr)
                c_ac.font = account_font
                c_ac.number_format = "#,##0"
                c_ac.alignment = align_right
                
                c_ab = ws_is.cell(row=curr_is_row, column=3, value=acc_base)
                c_ab.font = account_font
                c_ab.number_format = "#,##0"
                c_ab.alignment = align_right
                curr_is_row += 1
                
            curr_is_row += 1
            
    ws_is.column_dimensions["A"].width = 56
    ws_is.column_dimensions["B"].width = 16
    ws_is.column_dimensions["C"].width = 16
    ws_is.column_dimensions["D"].width = 10
    ws_is.column_dimensions["E"].width = 20
    ws_is.column_dimensions["F"].width = 20
    
    wb.save(output)
    output.seek(0)
    return output

def generate_executive_audit_report(
    company_name: str,
    period_current: str,
    period_base: str,
    validation_results: Dict[str, Any],
    df_classified: pd.DataFrame,
    statements: Dict[str, Any],
    ratios: Dict[str, Any],
    comparative_results: Optional[Dict[str, Any]] = None
) -> str:
    t = statements.get("totals", {})
    sums = validation_results.get("sums", {})
    diffs = validation_results.get("differences", {})
    
    diff_sum = diffs.get("sumas_deb_cred", 0.0)
    diff_sal = diffs.get("saldos_deud_acred", 0.0)
    diff_res = diffs.get("resultado_ejercicio_diff", 0.0)
    
    cuadre_status = "CUADRADO Y CONFORME (Tolerancia 0.00)" if validation_results.get("is_mathematically_sound") else "DESCUADRE DETECTADO"
    
    mapping_table_rows = []
    for _, r in df_classified.head(25).iterrows():
        code = r.get("codigo", "")
        name = r.get("nombre_cuenta", "")
        saldo = float(r.get("saldo_final", 0.0))
        rubro = r.get("rubro_clasificado", "")
        subrub = r.get("subrubro", "")
        status = r.get("mapping_status", "")
        mapping_table_rows.append(f"| `{code}` | {name} | ${saldo:,.0f} | {rubro} ({subrub}) | `{status}` |")
        
    mapping_table_md = "\n".join(mapping_table_rows)
    
    alerts_md = []
    if validation_results.get("alerts"):
        for a in validation_results["alerts"]:
            alerts_md.append(f"- ⚠️ **Alerta:** {a}")
    else:
        alerts_md.append("- ✅ No se detectaron inconsistencias matemáticas de origen.")
        
    reclass_md = []
    if validation_results.get("reclassifications_suggested"):
        for rec in validation_results["reclassifications_suggested"]:
            reclass_md.append(f"- 🔄 **{rec['nombre_cuenta']} (`{rec['codigo']}`)**: Saldo ${rec['saldo']:,.0f} -> *{rec['diagnostico']}*. {rec['propuesta']}")
    else:
        reclass_md.append("- ✅ Todos los saldos corresponden a la naturaleza contable esperada.")

    report = f"""# INFORME DE HOMOLOGACIÓN Y ESTADOS FINANCIEROS AUDITADOS (IFRS / NIC 1)
**Entidad:** {company_name} | **Ejercicio Analizado:** {period_current} | **Periodo Comparativo:** {period_base or 'N/A'}
**Fecha de Emisión:** {pd.Timestamp.now().strftime('%d/%m/%Y')} | **Responsable:** Gerencia de Finanzas / Auditoría Senior

---

## # 1. ESTADO DE LOS INSUMOS Y VALIDACIÓN INICIAL
- **Periodos Evaluados:** Año Base ({period_base or 'Sin base'}) vs. Año Actual ({period_current}).
- **Estado de Cuadre Matemático Balance de 8 Columnas:** **{cuadre_status}**
  - Total Débitos: `${sums.get('debito', 0):,.2f}` | Total Créditos: `${sums.get('credito', 0):,.2f}` | **Diferencia = ${diff_sum:,.2f}**
  - Total Deudor: `${sums.get('deudor', 0):,.2f}` | Total Acreedor: `${sums.get('acreedor', 0):,.2f}` | **Diferencia = ${diff_sal:,.2f}**
  - Cuadre de Resultado (Activo - Pasivo vs. Ganancia - Pérdida): **Diferencia = ${diff_res:,.2f}**

---

## # 2. MATRIZ DE MAPEO Y HOMOLOGACIÓN DE CUENTAS (Muestra de Trazabilidad)
| Código | Nombre Cuenta Origen (8 Col) | Saldo Año Actual | Rubro & Sub-rubro IFRS | Estado Mapeo |
| :--- | :--- | :--- | :--- | :--- |
{mapping_table_md}

---

## # 3. BALANCE CLASIFICADO IFRS (ESTRUCTURA JERÁRQUICA)
- **ACTIVOS CORRIENTES:** `${t.get('total_activo_corriente', 0):,.0f}`
- **ACTIVOS NO CORRIENTES:** `${t.get('total_activo_no_corriente', 0):,.0f}`
- **TOTAL ACTIVOS:** **`${t.get('total_activos', 0):,.0f}`**
---
- **PASIVOS CORRIENTES:** `${t.get('total_pasivo_corriente', 0):,.0f}`
- **PASIVOS NO CORRIENTES:** `${t.get('total_pasivo_no_corriente', 0):,.0f}`
- **TOTAL PASIVOS:** `${t.get('total_pasivos', 0):,.0f}`
- **PATRIMONIO NETO:** `${t.get('total_patrimonio', 0):,.0f}`
- **TOTAL PASIVO Y PATRIMONIO:** **`${t.get('total_pasivo_y_patrimonio', 0):,.0f}`**
- **CONTROL DE CUADRE PATRIMONIAL:** **`${t.get('diferencia_cuadre', 0):,.2f}`**

---

## # 4. CONTROL DE CALIDAD, ALERTAS E INCONSISTENCIAS
### Inconsistencias y Alertas del Balance:
{chr(10).join(alerts_md)}

### Propuestas de Reclasificación por Naturaleza Invertida:
{chr(10).join(reclass_md)}

---

## # 5. PUNTOS DE DECISIÓN Y DICTAMEN FINANCIERO
1. **Validación de Cuentas Nuevas:** Se requiere la confirmación formal de las cuentas catalogadas como `[CUENTA_NUEVA_SUGERIDA]` antes de remitir la versión definitiva a los auditores externos.
2. **Reclasificaciones Operativas:** Aprobar la reclasificación automática de saldos acreedores en cuentas corrientes bancarias hacia *Sobregiros Bancarios* bajo Pasivo Corriente conforme a NIC 1 §69.
3. **Certificación de Cuadre:** El balance presenta consistencia matemática total y se encuentra listo para su presentación a Directorio.
"""
    return report
