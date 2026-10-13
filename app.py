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
from typing import Any, Dict, List, Optional, Tuple

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
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
LOCAL_2024 = os.path.join(PROJECT_ROOT, "balance_base_2024.xlsx")
LOCAL_2025 = os.path.join(PROJECT_ROOT, "balance_base_2025.xlsx")
USER_FOLDER = r"C:\Users\omarp\OneDrive - himmel\EEFF 2025 Himmel"
DEFAULT_2024_PATH = os.path.join(USER_FOLDER, "Balance 2024 Himmel Spa 16 04 2025.xlsx")
DEFAULT_2025_PATH = os.path.join(USER_FOLDER, "Balance Himmel 2025.xlsx")

def get_himmel_default_paths():
    p24 = LOCAL_2024 if os.path.exists(LOCAL_2024) else (DEFAULT_2024_PATH if os.path.exists(DEFAULT_2024_PATH) else None)
    p25 = LOCAL_2025 if os.path.exists(LOCAL_2025) else (DEFAULT_2025_PATH if os.path.exists(DEFAULT_2025_PATH) else None)
    return p24, p25

def clean_code(c: Any) -> str:
    if pd.isna(c) or c is None:
        return ""
    s = str(c).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.lstrip("'").strip()

def load_default_himmel_data():
    p24, p25 = get_himmel_default_paths()
    if p24 and p25:
        df_24, _ = parse_8_column_balance(p24)
        df_25, _ = parse_8_column_balance(p25)
        st.session_state.df_base_8col = df_24
        st.session_state.df_current_8col = df_25
        st.session_state.df_base_classified = map_and_classify_dataframe(df_24)
        st.session_state.df_current_classified = map_and_classify_dataframe(df_25)
        st.session_state.current_period = "Dic 2025"
        st.session_state.base_period = "Dic 2024"
    else:
        demo_base, demo_curr = create_sample_balances()
        st.session_state.df_base_8col = demo_base
        st.session_state.df_current_8col = demo_curr
        st.session_state.df_base_classified = map_and_classify_dataframe(demo_base)
        st.session_state.df_current_classified = map_and_classify_dataframe(demo_curr)

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

# Auto-load real files if available on disk (local or cloud)
if st.session_state.df_current_8col is None:
    load_default_himmel_data()

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
    if st.button("🔄 Recargar Balances Oficiales", use_container_width=True):
        load_default_himmel_data()
        st.success("Balances oficiales recargados exitosamente.")
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
        base_lookup[clean_code(r.get("codigo", ""))] = float(r.get("saldo_final", 0.0))

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
                        sub_tot_base = sum(base_lookup.get(clean_code(c), 0.0) for c in sub_df["codigo"])
                        sub_nota = sub_df.iloc[0].get("nota_auditada", "")
                        
                        s_c1, s_c2, s_c3, s_c4, s_c5, s_c6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
                        with s_c1: st.markdown(f"&nbsp;&nbsp;*{sub_name}*")
                        with s_c4: st.markdown(f"<div style='text-align: center; color: #555;'>{sub_nota}</div>", unsafe_allow_html=True)
                        with s_c5: st.markdown(f"<div style='text-align: right; font-weight: bold;'>{sub_tot_curr:,.0f}</div>", unsafe_allow_html=True)
                        with s_c6: st.markdown(f"<div style='text-align: right; font-weight: bold;'>{sub_tot_base:,.0f}</div>", unsafe_allow_html=True)
                        
                        for _, acc_r in sub_df.iterrows():
                            acc_code = clean_code(acc_r.get("codigo", ""))
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
            tot_sec_base = sum(base_lookup.get(clean_code(c), 0.0) for c in sec_df["codigo"])
            
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
        is_df = df_curr[(df_curr["sheet_target"] == "EERR") | (df_curr["seccion"] == "ESTADO DE RESULTADOS") | df_curr["codigo"].str.startswith("4") | df_curr["codigo"].str.startswith("5")]
        
        def get_eerr_line_df(df, key):
            if df is None or df.empty:
                return pd.DataFrame()
            if key == "ingresos":
                return df[df["rubro_clasificado"].str.contains("ingreso", case=False, na=False) | (df["rubro_clasificado"].str.contains("venta", case=False, na=False) & ~df["rubro_clasificado"].str.contains("costo", case=False, na=False))]
            elif key == "costos":
                return df[df["rubro_clasificado"].str.contains("costo de venta", case=False, na=False) | (df["rubro_clasificado"].str.contains("costo", case=False, na=False) & ~df["rubro_clasificado"].str.contains("financier", case=False, na=False))]
            elif key == "admin":
                return df[df["rubro_clasificado"].str.contains("administra", case=False, na=False)]
            elif key == "financieros":
                return df[df["rubro_clasificado"].str.contains("financier", case=False, na=False) | df["rubro_clasificado"].str.contains("interes", case=False, na=False)]
            elif key == "cambio":
                return df[df["rubro_clasificado"].str.contains("cambio", case=False, na=False)]
            elif key == "monetaria":
                return df[df["rubro_clasificado"].str.contains("monetaria", case=False, na=False) | df["rubro_clasificado"].str.contains("reajuste", case=False, na=False)]
            return df.iloc[0:0]

        eerr_sections = [
            ("ingresos", "Ingresos de Actividades Ordinarias", 1),
            ("costos", "Costo de Ventas", -1),
            ("bruta", "GANANCIA BRUTA", 0),
            ("admin", "Gastos de Administración y Ventas", -1),
            ("operacional", "GANANCIAS (PÉRDIDAS) OPERACIONALES", 0),
            ("financieros", "Costos Financieros", -1),
            ("cambio", "Diferencias de Cambio", -1),
            ("monetaria", "Corrección Monetaria", -1),
            ("ejercicio", "GANANCIA (PÉRDIDA) DEL EJERCICIO", 0)
        ]
        
        tot_ing_curr = get_eerr_line_df(is_df, "ingresos")["saldo_final"].sum()
        tot_cos_curr = get_eerr_line_df(is_df, "costos")["saldo_final"].sum()
        tot_adm_curr = get_eerr_line_df(is_df, "admin")["saldo_final"].sum()
        tot_fin_curr = get_eerr_line_df(is_df, "financieros")["saldo_final"].sum()
        tot_tc_curr = get_eerr_line_df(is_df, "cambio")["saldo_final"].sum()
        tot_cm_curr = get_eerr_line_df(is_df, "monetaria")["saldo_final"].sum()

        tot_ing_base = sum(base_lookup.get(clean_code(c), 0.0) for c in get_eerr_line_df(is_df, "ingresos")["codigo"])
        tot_cos_base = sum(base_lookup.get(clean_code(c), 0.0) for c in get_eerr_line_df(is_df, "costos")["codigo"])
        tot_adm_base = sum(base_lookup.get(clean_code(c), 0.0) for c in get_eerr_line_df(is_df, "admin")["codigo"])
        tot_fin_base = sum(base_lookup.get(clean_code(c), 0.0) for c in get_eerr_line_df(is_df, "financieros")["codigo"])
        tot_tc_base = sum(base_lookup.get(clean_code(c), 0.0) for c in get_eerr_line_df(is_df, "cambio")["codigo"])
        tot_cm_base = sum(base_lookup.get(clean_code(c), 0.0) for c in get_eerr_line_df(is_df, "monetaria")["codigo"])

        # Also check base classified if available to capture full totals
        if st.session_state.df_base_classified is not None:
            df_b = st.session_state.df_base_classified
            is_b_df = df_b[(df_b["sheet_target"] == "EERR") | (df_b["seccion"] == "ESTADO DE RESULTADOS") | df_b["codigo"].str.startswith("4") | df_b["codigo"].str.startswith("5")]
            b_ing = get_eerr_line_df(is_b_df, "ingresos")["saldo_final"].sum()
            b_cos = get_eerr_line_df(is_b_df, "costos")["saldo_final"].sum()
            b_adm = get_eerr_line_df(is_b_df, "admin")["saldo_final"].sum()
            b_fin = get_eerr_line_df(is_b_df, "financieros")["saldo_final"].sum()
            b_tc = get_eerr_line_df(is_b_df, "cambio")["saldo_final"].sum()
            b_cm = get_eerr_line_df(is_b_df, "monetaria")["saldo_final"].sum()
            if b_ing > 0 or b_adm > 0:
                tot_ing_base, tot_cos_base, tot_adm_base, tot_fin_base, tot_tc_base, tot_cm_base = b_ing, b_cos, b_adm, b_fin, b_tc, b_cm

        for sec_key, rub_label, multiplier in eerr_sections:
            if multiplier == 0:
                if rub_label == "GANANCIA BRUTA":
                    val_c = tot_ing_curr - tot_cos_curr
                    val_b = tot_ing_base - tot_cos_base
                elif "OPERACIONALES" in rub_label:
                    val_c = (tot_ing_curr - tot_cos_curr) - tot_adm_curr
                    val_b = (tot_ing_base - tot_cos_base) - tot_adm_base
                else: # Net Income
                    val_c = (tot_ing_curr - tot_cos_curr) - tot_adm_curr - tot_fin_curr - tot_tc_curr - tot_cm_curr
                    val_b = (tot_ing_base - tot_cos_base) - tot_adm_base - tot_fin_base - tot_tc_base - tot_cm_base
                    
                st.divider()
                s_c1, s_c2, s_c3, s_c4, s_c5, s_c6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
                with s_c1: st.markdown(f"### **{rub_label}**")
                with s_c5: st.markdown(f"<div style='text-align: right; font-size: 1.25rem; font-weight: bold; color: #1F4E79;'>$ {val_c:,.0f}</div>", unsafe_allow_html=True)
                with s_c6: st.markdown(f"<div style='text-align: right; font-size: 1.25rem; font-weight: bold; color: #1F4E79;'>$ {val_b:,.0f}</div>", unsafe_allow_html=True)
                st.divider()
            else:
                sub_is_df = get_eerr_line_df(is_df, sec_key)
                sub_tot_curr = sub_is_df["saldo_final"].sum()
                sub_tot_base = sum(base_lookup.get(clean_code(c), 0.0) for c in sub_is_df["codigo"])
                
                s_c1, s_c2, s_c3, s_c4, s_c5, s_c6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
                with s_c1: st.markdown(f"##### **{rub_label}**")
                with s_c5: st.markdown(f"<div style='text-align: right; font-weight: bold;'>{(sub_tot_curr*multiplier):,.0f}</div>", unsafe_allow_html=True)
                with s_c6: st.markdown(f"<div style='text-align: right; font-weight: bold;'>{(sub_tot_base*multiplier):,.0f}</div>", unsafe_allow_html=True)
                
                for _, acc_r in sub_is_df.iterrows():
                    acc_code = clean_code(acc_r.get("codigo", ""))
                    acc_name = str(acc_r.get("nombre_cuenta", "")).strip()
                    acc_curr = float(acc_r.get("saldo_final", 0.0))
                    acc_base = base_lookup.get(acc_code, 0.0)
                    
                    a_c1, a_c2, a_c3, a_c4, a_c5, a_c6 = st.columns([5, 2, 2, 1, 2.5, 2.5])
                    with a_c1: st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`{acc_code}` {acc_name}")
                    with a_c2: st.markdown(f"<div style='text-align: right;'>{acc_curr:,.0f}</div>", unsafe_allow_html=True)
                    with a_c3: st.markdown(f"<div style='text-align: right;'>{acc_base:,.0f}</div>", unsafe_allow_html=True)
                st.write("")

# ==============================================================================
# TAB 3: CARGA DE BALANCES 8 COLUMNAS
# ==============================================================================
with tab3:
    st.subheader("📤 Carga de Balances Generales de 8 Columnas")
    st.caption("Carga los balances de 8 columnas en formato Excel (.xlsx, .xls) o CSV. El motor procesará y homologará automáticamente cada cuenta bajo IFRS / CMF.")
    
    col_u1, col_u2 = st.columns(2)
    
    with col_u1:
        st.markdown("### 🟢 1. Balance Período Actual")
        p_act_input = st.text_input("Etiqueta Período Actual:", value=st.session_state.current_period, key="p_act_inp")
        file_curr = st.file_uploader("Subir Balance Actual (ej: Agosto 2026):", type=["xlsx", "xls", "csv"], key="file_curr_inp")
        if file_curr is not None:
            df_curr_p, meta_c = parse_8_column_balance(file_curr)
            if meta_c["is_valid"]:
                st.session_state.df_current_8col = df_curr_p
                st.session_state.df_current_classified = map_and_classify_dataframe(df_curr_p)
                st.session_state.current_period = p_act_input
                st.success(f"✅ Balance Actual procesado exitosamente ({len(df_curr_p)} cuentas).")
                st.info(f"Hoja: `{meta_c.get('sheet_used')}` | Fila encabezado: `{meta_c.get('header_row_index')}`")
            else:
                st.error(f"❌ Error al procesar: {meta_c.get('error_message')}")
                
    with col_u2:
        st.markdown("### 🔵 2. Balance Período Base (Comparativo)")
        p_base_input = st.text_input("Etiqueta Período Base:", value=st.session_state.base_period, key="p_base_inp")
        file_base = st.file_uploader("Subir Balance Base (ej: Diciembre 2025):", type=["xlsx", "xls", "csv"], key="file_base_inp")
        if file_base is not None:
            df_base_p, meta_b = parse_8_column_balance(file_base)
            if meta_b["is_valid"]:
                st.session_state.df_base_8col = df_base_p
                st.session_state.df_base_classified = map_and_classify_dataframe(df_base_p)
                st.session_state.base_period = p_base_input
                st.success(f"✅ Balance Base procesado exitosamente ({len(df_base_p)} cuentas).")
                st.info(f"Hoja: `{meta_b.get('sheet_used')}` | Fila encabezado: `{meta_b.get('header_row_index')}`")
            else:
                st.error(f"❌ Error al procesar: {meta_b.get('error_message')}")
                
        p24, p25 = get_himmel_default_paths()
        if p25:
            if st.button("📌 Fijar Balance Himmel 2025 Oficial como Período Base", use_container_width=True):
                df_b25, _ = parse_8_column_balance(p25)
                st.session_state.df_base_8col = df_b25
                st.session_state.df_base_classified = map_and_classify_dataframe(df_b25)
                st.session_state.base_period = "Dic 2025"
                st.success("✅ Balance Dic 2025 cargado como Período Base.")
                st.rerun()

    st.divider()
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        if st.session_state.df_current_classified is not None:
            tot_act = st.session_state.df_current_classified[st.session_state.df_current_classified["seccion"] == "ACTIVOS"]["saldo_final"].sum()
            tot_pas_pat = st.session_state.df_current_classified[st.session_state.df_current_classified["seccion"] == "PASIVOS Y PATRIMONIO"]["saldo_final"].sum()
            st.metric(f"Total Activos ({st.session_state.current_period})", f"${tot_act:,.0f}")
            st.metric(f"Total Pasivo + Pat. ({st.session_state.current_period})", f"${tot_pas_pat:,.0f}", delta=f"${(tot_act - tot_pas_pat):,.0f} cuadre")
    with m_col2:
        if st.session_state.df_base_classified is not None:
            tot_act_b = st.session_state.df_base_classified[st.session_state.df_base_classified["seccion"] == "ACTIVOS"]["saldo_final"].sum()
            tot_pas_pat_b = st.session_state.df_base_classified[st.session_state.df_base_classified["seccion"] == "PASIVOS Y PATRIMONIO"]["saldo_final"].sum()
            st.metric(f"Total Activos Base ({st.session_state.base_period})", f"${tot_act_b:,.0f}")
            st.metric(f"Total Pasivo + Pat. Base ({st.session_state.base_period})", f"${tot_pas_pat_b:,.0f}", delta=f"${(tot_act_b - tot_pas_pat_b):,.0f} cuadre")
            
    if st.session_state.df_current_8col is not None:
        st.markdown(f"#### Vista Previa del Balance de Origen ({st.session_state.current_period}):")
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
