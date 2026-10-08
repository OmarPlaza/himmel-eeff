"""
CFO Financial Intelligence & IFRS Automation Suite
Enterprise-grade Streamlit application for 8-column trial balance ingestion,
deterministic auditing, interactive IFRS classification, multi-period comparison,
and executive financial statement presentation matching official audited formats (Himmel SpA).
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import os
import sys

# Add project root to sys.path
sys.path.append(os.path.dirname(__file__))

from core.database import init_db, get_connection
from core.parser import parse_8_column_balance
from core.validator import validate_8_column_balance
from core.classifier import (
    map_and_classify_dataframe,
    build_classified_financial_statements,
    get_all_rubros_list,
    save_custom_mapping
)
from core.comparator import compare_classified_balances
from core.ratios import calculate_financial_ratios
from core.exporter import generate_audited_detailed_excel, generate_executive_audit_report
from core.sample_data import create_sample_balances

# Page configuration
st.set_page_config(
    page_title="CFO Suite | Homologación & Estados Financieros IFRS",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Corporate CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1F4E79;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #5A6B7C;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #F8FAFC 0%, #EDF2F7 100%);
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1.2rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
    }
    .audited-header {
        background-color: #FFFFFF;
        border: 1px solid #000000;
        padding: 8px;
        font-weight: bold;
        text-align: center;
        letter-spacing: 2px;
        margin-top: 15px;
        margin-bottom: 10px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        white-space: pre-wrap;
        background-color: #F1F5F9;
        border-radius: 8px 8px 0px 0px;
        font-weight: 600;
        padding: 0 16px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1F4E79 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# Default paths for Himmel SpA
USER_FOLDER = r"C:\Users\omarp\OneDrive - himmel\EEFF 2025 Himmel"
DEFAULT_2024_PATH = os.path.join(USER_FOLDER, "Balance 2024 Himmel Spa 16 04 2025.xlsx")
DEFAULT_2025_PATH = os.path.join(USER_FOLDER, "Balance Himmel 2025.xlsx")

# Initialize Session State
if "df_current_8col" not in st.session_state:
    st.session_state.df_current_8col = None
if "df_base_8col" not in st.session_state:
    st.session_state.df_base_8col = None
if "df_current_classified" not in st.session_state:
    st.session_state.df_current_classified = None
if "df_base_classified" not in st.session_state:
    st.session_state.df_base_classified = None
if "company_name" not in st.session_state:
    st.session_state.company_name = "HIMMEL SPA"
if "current_period" not in st.session_state:
    st.session_state.current_period = "Agos 2026"
if "base_period" not in st.session_state:
    st.session_state.base_period = "Dic 2025"

# Auto-load real files if available on disk
if st.session_state.df_current_8col is None:
    if os.path.exists(DEFAULT_2025_PATH) and os.path.exists(DEFAULT_2024_PATH):
        df_24, _ = parse_8_column_balance(DEFAULT_2024_PATH)
        df_25, _ = parse_8_column_balance(DEFAULT_2025_PATH)
        st.session_state.df_base_8col = df_24
        st.session_state.df_current_8col = df_25
        st.session_state.df_base_classified = map_and_classify_dataframe(df_24)
        st.session_state.df_current_classified = map_and_classify_dataframe(df_25)
    else:
        demo_base, demo_curr = create_sample_balances()
        st.session_state.df_base_8col = demo_base
        st.session_state.df_current_8col = demo_curr
        st.session_state.df_base_classified = map_and_classify_dataframe(demo_base)
        st.session_state.df_current_classified = map_and_classify_dataframe(demo_curr)

# Sidebar: Company & Configuration
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/accounting.png", width=64)
    st.title("CFO Suite IFRS")
    st.caption("Sistema de Homologación & Auditoría Financiera")
    st.divider()
    
    st.subheader("🏢 Empresa y Ejercicios")
    st.session_state.company_name = st.text_input("Razón Social:", value=st.session_state.company_name)
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.session_state.current_period = st.text_input("Período Actual:", value=st.session_state.current_period)
    with col_p2:
        st.session_state.base_period = st.text_input("Período Base:", value=st.session_state.base_period)
        
    st.divider()
    
    st.subheader("⚡ Acciones Rápidas")
    if st.button("🔄 Recargar Balances de Himmel", use_container_width=True):
        if os.path.exists(DEFAULT_2025_PATH) and os.path.exists(DEFAULT_2024_PATH):
            df_24, _ = parse_8_column_balance(DEFAULT_2024_PATH)
            df_25, _ = parse_8_column_balance(DEFAULT_2025_PATH)
            st.session_state.df_base_8col = df_24
            st.session_state.df_current_8col = df_25
            st.session_state.df_base_classified = map_and_classify_dataframe(df_24)
            st.session_state.df_current_classified = map_and_classify_dataframe(df_25)
            st.success("Balances recargados.")
            st.rerun()
            
    if st.button("🗑️ Limpiar Datos", use_container_width=True):
        st.session_state.df_current_8col = None
        st.session_state.df_base_8col = None
        st.session_state.df_current_classified = None
        st.session_state.df_base_classified = None
        st.rerun()
        
    st.divider()
    st.caption("Normativa: IFRS Full / NIC 1 / CMF")

# Main Header
st.markdown(f"<div class='main-header'>💼 Homologación & Estados Financieros Auditados | {st.session_state.company_name}</div>", unsafe_allow_html=True)
st.markdown(f"<div class='sub-header'>Conversión automática de Balance de 8 Columnas a Estados Financieros Clasificados (Balance Comparativo & EERR)</div>", unsafe_allow_html=True)

# Tabs
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📑 1. Balance Clasificado Auditado",
    "📊 2. Estado de Resultados (EERR)",
    "📤 3. Carga de Balance (8 Columnas)",
    "🛠️ 4. Matriz de Mapeo & Corrección",
    "📈 5. Comparador de Cuentas",
    "💾 6. Descargar Excel Formateado"
])

# Base lookup dictionary
base_lookup = {}
if st.session_state.df_base_classified is not None:
    for _, r in st.session_state.df_base_classified.iterrows():
        base_lookup[str(r.get("codigo", "")).strip()] = float(r.get("saldo_final", 0.0))

# ==============================================================================
# TAB 1: BALANCE CLASIFICADO AUDITADO (EXACTO FORMATO IMAGEN DERECHA)
# ==============================================================================
with tab1:
    if st.session_state.df_current_classified is not None:
        p_clean = st.session_state.current_period.replace("Dic ", "").replace("Agos ", "").replace("Ene ", "")
        st.markdown(f"<h3 style='text-align: center; margin-bottom: 0px;'>{st.session_state.company_name}</h3>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; color: #666;'>Balance General entre el 01 de Enero de {p_clean} al período {st.session_state.current_period}</p>", unsafe_allow_html=True)
        
        # Download Excel
        excel_bytes = generate_audited_detailed_excel(
            st.session_state.df_current_classified,
            df_base_classified=st.session_state.df_base_classified,
            company_name=st.session_state.company_name,
            period_current=st.session_state.current_period,
            period_base=st.session_state.base_period
        )
        st.download_button(
            label="📥 Descargar Balance y EERR en Excel (.xlsx)",
            data=excel_bytes.getvalue(),
            file_name=f"EEFF_Auditados_{st.session_state.company_name.replace(' ', '_')}_{st.session_state.current_period}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        df_curr = st.session_state.df_current_classified
        
        sections = [
            ("ACTIVOS", "A C T I V O S", ["ACTIVOS CORRIENTES", "ACTIVOS NO CORRIENTES"]),
            ("PASIVOS Y PATRIMONIO", "P A S I V O S   Y   P A T R I M O N I O", ["PASIVOS CORRIENTES", "PASIVOS NO CORRIENTES", "PATRIMONIO"])
        ]
        
        for sec_id, sec_title, categories in sections:
            st.markdown(f"<div class='audited-header'>{sec_title}</div>", unsafe_allow_html=True)
            
            # Header Columns
            h_col1, h_col2, h_col3, h_col4, h_col5, h_col6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
            with h_col1: st.markdown(f"**{categories[0]}**")
            with h_col2: st.markdown(f"<div style='text-align: right;'><b>{st.session_state.current_period}</b></div>", unsafe_allow_html=True)
            with h_col3: st.markdown(f"<div style='text-align: right;'><b>{st.session_state.base_period}</b></div>", unsafe_allow_html=True)
            with h_col4: st.markdown("<div style='text-align: center;'><b>Nota</b></div>", unsafe_allow_html=True)
            with h_col5: st.markdown(f"<div style='text-align: right;'><b>$ {st.session_state.current_period}</b></div>", unsafe_allow_html=True)
            with h_col6: st.markdown(f"<div style='text-align: right;'><b>$ {st.session_state.base_period}</b></div>", unsafe_allow_html=True)
            st.divider()
            
            sec_df = df_curr[(df_curr["seccion"] == sec_id) & (df_curr["sheet_target"] != "EERR")]
            
            for cat_key in categories:
                cat_df = sec_df[sec_df["categoria_ifrs"] == cat_key]
                if cat_df.empty:
                    continue
                    
                if cat_key != categories[0]:
                    st.markdown(f"#### **{cat_key}**")
                    
                rubros = cat_df["rubro_clasificado"].unique()
                for rub_name in rubros:
                    rub_df = cat_df[cat_df["rubro_clasificado"] == rub_name]
                    st.markdown(f"##### **{rub_name}**")
                    
                    subrubros = rub_df["subrubro"].unique()
                    for sub_name in subrubros:
                        sub_df = rub_df[rub_df["subrubro"] == sub_name]
                        sub_tot_curr = sub_df["saldo_final"].sum()
                        sub_tot_base = sum(base_lookup.get(str(c).strip(), 0.0) for c in sub_df["codigo"])
                        sub_nota = sub_df.iloc[0].get("nota_auditada", "")
                        
                        s_c1, s_c2, s_c3, s_c4, s_c5, s_c6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
                        with s_c1: st.markdown(f"&nbsp;&nbsp;*{sub_name}*")
                        with s_c4: st.markdown(f"<div style='text-align: center; color: #555;'>{sub_nota}</div>", unsafe_allow_html=True)
                        with s_c5: st.markdown(f"<div style='text-align: right; font-weight: bold;'>{sub_tot_curr:,.0f}</div>", unsafe_allow_html=True)
                        with s_c6: st.markdown(f"<div style='text-align: right; font-weight: bold;'>{sub_tot_base:,.0f}</div>", unsafe_allow_html=True)
                        
                        for _, acc_r in sub_df.iterrows():
                            acc_code = str(acc_r.get("codigo", "")).strip()
                            acc_name = str(acc_r.get("nombre_cuenta", "")).strip()
                            acc_curr = float(acc_r.get("saldo_final", 0.0))
                            acc_base = base_lookup.get(acc_code, 0.0)
                            acc_nota = str(acc_r.get("nota_auditada", ""))
                            
                            a_c1, a_c2, a_c3, a_c4, a_c5, a_c6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
                            with a_c1: st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`{acc_code}` {acc_name}")
                            with a_c2: st.markdown(f"<div style='text-align: right;'>{acc_curr:,.0f}</div>", unsafe_allow_html=True)
                            with a_c3: st.markdown(f"<div style='text-align: right;'>{acc_base:,.0f}</div>", unsafe_allow_html=True)
                            with a_c4: 
                                if acc_nota and acc_nota != sub_nota and acc_nota != "-":
                                    st.markdown(f"<div style='text-align: center; font-size: 0.85rem;'>{acc_nota}</div>", unsafe_allow_html=True)
                        st.write("")
                        
            tot_sec_curr = sec_df["saldo_final"].sum()
            tot_sec_base = sum(base_lookup.get(str(c).strip(), 0.0) for c in sec_df["codigo"])
            
            st.divider()
            t_c1, t_c2, t_c3, t_c4, t_c5, t_c6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
            with t_c1: st.markdown(f"### **TOTAL {sec_id}**")
            with t_c5: st.markdown(f"<div style='text-align: right; font-size: 1.3rem; font-weight: bold; color: #1F4E79;'>$ {tot_sec_curr:,.0f}</div>", unsafe_allow_html=True)
            with t_c6: st.markdown(f"<div style='text-align: right; font-size: 1.3rem; font-weight: bold; color: #1F4E79;'>$ {tot_sec_base:,.0f}</div>", unsafe_allow_html=True)
            st.write("")

# ==============================================================================
# TAB 2: ESTADO DE RESULTADOS (EERR)
# ==============================================================================
with tab2:
    if st.session_state.df_current_classified is not None:
        p_clean = st.session_state.current_period.replace("Dic ", "").replace("Agos ", "").replace("Ene ", "")
        st.markdown(f"<h3 style='text-align: center; margin-bottom: 0px;'>{st.session_state.company_name}</h3>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; color: #666;'>Estados de Resultados por función del 01 de enero de {p_clean} al período {st.session_state.current_period}</p>", unsafe_allow_html=True)
        
        st.markdown("<div class='audited-header'>E S T A D O   D E   R E S U L T A D O S</div>", unsafe_allow_html=True)
        
        # Header Columns
        h_col1, h_col2, h_col3, h_col4, h_col5, h_col6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
        with h_col1: st.markdown("**LÍNEA / CUENTA**")
        with h_col2: st.markdown(f"<div style='text-align: right;'><b>{st.session_state.current_period}</b></div>", unsafe_allow_html=True)
        with h_col3: st.markdown(f"<div style='text-align: right;'><b>{st.session_state.base_period}</b></div>", unsafe_allow_html=True)
        with h_col4: st.markdown("<div style='text-align: center;'><b>Nota</b></div>", unsafe_allow_html=True)
        with h_col5: st.markdown(f"<div style='text-align: right;'><b>$ {st.session_state.current_period}</b></div>", unsafe_allow_html=True)
        with h_col6: st.markdown(f"<div style='text-align: right;'><b>$ {st.session_state.base_period}</b></div>", unsafe_allow_html=True)
        st.divider()
        
        df_curr = st.session_state.df_current_classified
        is_df = df_curr[(df_curr["sheet_target"] == "EERR") | (df_curr["seccion"] == "ESTADO DE RESULTADOS")]
        
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
        
        tot_ing_curr = is_df[is_df["rubro_clasificado"] == "Ingresos de Actividades Ordinarias"]["saldo_final"].sum()
        tot_cos_curr = is_df[is_df["rubro_clasificado"] == "Costo de Ventas"]["saldo_final"].sum()
        tot_adm_curr = is_df[is_df["rubro_clasificado"] == "Gastos de Administración"]["saldo_final"].sum()
        tot_fin_curr = is_df[is_df["rubro_clasificado"] == "Costos Financieros"]["saldo_final"].sum()
        tot_tc_curr = is_df[is_df["rubro_clasificado"] == "Diferencias de cambios por conversión"]["saldo_final"].sum()
        tot_cm_curr = is_df[is_df["rubro_clasificado"] == "Corrección Monetaria"]["saldo_final"].sum()
        
        for rub_key, rub_label, multiplier in eerr_sections:
            if multiplier == 0:
                if rub_label == "GANANCIA BRUTA":
                    val_c = tot_ing_curr - tot_cos_curr
                elif "OPERACIONALES" in rub_label:
                    val_c = (tot_ing_curr - tot_cos_curr) - tot_adm_curr
                else: # Net Income
                    val_c = (tot_ing_curr - tot_cos_curr) - tot_adm_curr - tot_fin_curr - tot_tc_curr - tot_cm_curr
                    
                st.divider()
                s_c1, s_c2, s_c3, s_c4, s_c5, s_c6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
                with s_c1: st.markdown(f"### **{rub_label}**")
                with s_c5: st.markdown(f"<div style='text-align: right; font-size: 1.25rem; font-weight: bold; color: #1F4E79;'>$ {val_c:,.0f}</div>", unsafe_allow_html=True)
                st.divider()
            else:
                sub_is_df = is_df[is_df["rubro_clasificado"] == rub_key]
                sub_tot_curr = sub_is_df["saldo_final"].sum()
                sub_tot_base = sum(base_lookup.get(str(c).strip(), 0.0) for c in sub_is_df["codigo"])
                
                s_c1, s_c2, s_c3, s_c4, s_c5, s_c6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
                with s_c1: st.markdown(f"##### **{rub_label}**")
                with s_c5: st.markdown(f"<div style='text-align: right; font-weight: bold;'>{(sub_tot_curr*multiplier):,.0f}</div>", unsafe_allow_html=True)
                with s_c6: st.markdown(f"<div style='text-align: right; font-weight: bold;'>{(sub_tot_base*multiplier):,.0f}</div>", unsafe_allow_html=True)
                
                for _, acc_r in sub_is_df.iterrows():
                    acc_code = str(acc_r.get("codigo", "")).strip()
                    acc_name = str(acc_r.get("nombre_cuenta", "")).strip()
                    acc_curr = float(acc_r.get("saldo_final", 0.0))
                    acc_base = base_lookup.get(acc_code, 0.0)
                    
                    a_c1, a_c2, a_c3, a_c4, a_c5, a_c6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
                    with a_c1: st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`{acc_code}` {acc_name}")
                    with a_c2: st.markdown(f"<div style='text-align: right;'>{acc_curr:,.0f}</div>", unsafe_allow_html=True)
                    with a_c3: st.markdown(f"<div style='text-align: right;'>{acc_base:,.0f}</div>", unsafe_allow_html=True)
                st.write("")

# ==============================================================================
# TAB 3: CARGA DE BALANCE 8 COLUMNAS
# ==============================================================================
with tab3:
    st.subheader("📤 Cargar Balance General de 8 Columnas")
    st.caption("Arrastra tu archivo Excel de 8 columnas. El sistema extraerá automáticamente todas las columnas contables y las homologará.")
    
    uploaded_file = st.file_uploader("Selecciona el Balance de 8 Columnas (.xlsx / .csv):", type=["xlsx", "xls", "csv"])
    if uploaded_file is not None:
        df_parsed, meta = parse_8_column_balance(uploaded_file)
        if meta["is_valid"]:
            st.session_state.df_current_8col = df_parsed
            st.session_state.df_current_classified = map_and_classify_dataframe(df_parsed)
            st.success(f"✅ Balance cargado con éxito ({len(df_parsed)} cuentas procesadas).")
        else:
            st.error(f"Error al procesar: {meta.get('error_message')}")
            
    if st.session_state.df_current_8col is not None:
        st.markdown("#### Vista Previa del Balance de Origen (8 Columnas):")
        st.dataframe(
            st.session_state.df_current_8col.style.format({
                "debito": "{:,.0f}", "credito": "{:,.0f}",
                "deudor": "{:,.0f}", "acreedor": "{:,.0f}",
                "activo": "{:,.0f}", "pasivo": "{:,.0f}",
                "perdida": "{:,.0f}", "ganancia": "{:,.0f}"
            }),
            use_container_width=True
        )

# ==============================================================================
# TAB 4: MATRIZ DE MAPEO & CORRECCIÓN
# ==============================================================================
with tab4:
    st.subheader("🛠️ Corregir / Reclasificar Cuentas")
    st.caption("Como Gerente de Finanzas, puedes ajustar el rubro clasificado y el sub-rubro de cualquier cuenta. Los cambios se reflejarán instantáneamente.")
    
    if st.session_state.df_current_classified is not None:
        all_rubros = get_all_rubros_list()
        
        col_search, col_filter = st.columns([2, 1])
        with col_search:
            q = st.text_input("Buscar cuenta por código o nombre:", "")
        with col_filter:
            status = st.selectbox("Estado de Mapeo:", ["TODOS", "HOMOLOGADO", "CUENTA_NUEVA_SUGERIDA", "IDÉNTICO"])
            
        df_edit = st.session_state.df_current_classified.copy()
        if q:
            df_edit = df_edit[df_edit["nombre_cuenta"].str.contains(q, case=False, na=False) | df_edit["codigo"].str.contains(q, case=False, na=False)]
        if status != "TODOS":
            df_edit = df_edit[df_edit["mapping_status"] == status]
            
        edited = st.data_editor(
            df_edit[["codigo", "nombre_cuenta", "saldo_final", "rubro_clasificado", "subrubro", "nota_auditada", "mapping_status"]],
            column_config={
                "codigo": st.column_config.TextColumn("Código", disabled=True),
                "nombre_cuenta": st.column_config.TextColumn("Nombre Cuenta", disabled=True),
                "saldo_final": st.column_config.NumberColumn("Saldo ($)", format="$%,.0f", disabled=True),
                "rubro_clasificado": st.column_config.SelectboxColumn("Rubro Clasificado", options=all_rubros, required=True),
                "subrubro": st.column_config.TextColumn("Sub-Rubro / Detalle"),
                "nota_auditada": st.column_config.TextColumn("Nota"),
                "mapping_status": st.column_config.TextColumn("Estado", disabled=True)
            },
            use_container_width=True,
            key="matrix_editor"
        )
        
        if st.button("💾 Guardar y Aplicar Clasificaciones", type="primary"):
            for _, r in edited.iterrows():
                c_code = r["codigo"]
                new_rub = r["rubro_clasificado"]
                new_sub = r["subrubro"]
                new_nota = r["nota_auditada"]
                idx = st.session_state.df_current_classified[st.session_state.df_current_classified["codigo"] == c_code].index
                if len(idx) > 0:
                    st.session_state.df_current_classified.loc[idx, "rubro_clasificado"] = new_rub
                    st.session_state.df_current_classified.loc[idx, "subrubro"] = new_sub
                    st.session_state.df_current_classified.loc[idx, "nota_auditada"] = new_nota
                    save_custom_mapping(c_code, st.session_state.df_current_classified.loc[idx[0], "nombre_cuenta"], new_rub, new_sub, new_nota)
            st.success("✅ Clasificaciones actualizadas y guardadas.")
            st.rerun()

# ==============================================================================
# TAB 5: COMPARADOR DE CUENTAS
# ==============================================================================
with tab5:
    st.subheader(f"📈 Comparación Multiperíodo: {st.session_state.base_period} vs {st.session_state.current_period}")
    if st.session_state.df_base_classified is not None and st.session_state.df_current_classified is not None:
        comp = compare_classified_balances(st.session_state.df_base_classified, st.session_state.df_current_classified)
        
        st.markdown("#### Detalle Comparativo Cuenta por Cuenta:")
        st.dataframe(
            comp["cuentas_comparativo_df"].style.format({
                "saldo_base": "{:,.0f}",
                "saldo_actual": "{:,.0f}",
                "variacion_nominal": "{:,.0f}",
                "variacion_pct": "{:+.1f}%"
            }),
            use_container_width=True,
            hide_index=True
        )

# ==============================================================================
# TAB 6: DESCARGAS
# ==============================================================================
with tab6:
    st.subheader("💾 Centro de Descarga de Planillas y Reportes")
    if st.session_state.df_current_classified is not None:
        excel_bytes = generate_audited_detailed_excel(
            st.session_state.df_current_classified,
            df_base_classified=st.session_state.df_base_classified,
            company_name=st.session_state.company_name,
            period_current=st.session_state.current_period,
            period_base=st.session_state.base_period
        )
        
        st.download_button(
            label="📥 Descargar Estados Financieros Completos (Balance + EERR) en Excel (.xlsx)",
            data=excel_bytes.getvalue(),
            file_name=f"EEFF_{st.session_state.company_name.replace(' ', '_')}_{st.session_state.current_period}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
