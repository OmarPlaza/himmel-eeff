"""
Database layer for CFO Financial Intelligence Suite
Provides SQLite persistence for companies, periods, balances, and mapping rules.
"""
import sqlite3
import json
import os
from datetime import datetime
from typing import List, Dict, Any, Optional

DB_PATH = os.path.join(os.path.dirname(__file__), "financial_system.db")

def get_connection(db_path: str = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def init_db(db_path: str = DB_PATH):
    conn = get_connection(db_path)
    cursor = conn.cursor()
    
    # Companies
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        tax_id TEXT,
        currency TEXT DEFAULT 'CLP',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)
    
    # Periods / Exercises
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS periods (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        year INTEGER NOT NULL,
        period_name TEXT NOT NULL,
        is_base_year INTEGER DEFAULT 0,
        status TEXT DEFAULT 'DRAFT',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (company_id) REFERENCES companies(id),
        UNIQUE(company_id, year, period_name)
    );
    """)
    
    # Mapping Rules (Master & Learned)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS mapping_rules (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER,
        account_code TEXT,
        account_pattern TEXT,
        account_name TEXT,
        rubro_clasificado TEXT NOT NULL,
        presentacion_auditada TEXT NOT NULL,
        categoria_ifrs TEXT NOT NULL,
        naturaleza TEXT NOT NULL,
        notes TEXT,
        confidence REAL DEFAULT 1.0,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (company_id) REFERENCES companies(id)
    );
    """)
    
    # Raw 8-Column Balance Accounts
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS balance_accounts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        period_id INTEGER NOT NULL,
        account_code TEXT NOT NULL,
        account_name TEXT NOT NULL,
        debito REAL DEFAULT 0,
        credito REAL DEFAULT 0,
        deudor REAL DEFAULT 0,
        acreedor REAL DEFAULT 0,
        activo REAL DEFAULT 0,
        pasivo REAL DEFAULT 0,
        perdida REAL DEFAULT 0,
        ganancia REAL DEFAULT 0,
        rubro_clasificado TEXT,
        presentacion_auditada TEXT,
        categoria_ifrs TEXT,
        mapping_status TEXT DEFAULT 'PENDIENTE',
        user_override INTEGER DEFAULT 0,
        reclassification_note TEXT,
        FOREIGN KEY (period_id) REFERENCES periods(id)
    );
    """)
    
    # Classification & Audit Logs
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS audit_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        period_id INTEGER NOT NULL,
        action TEXT NOT NULL,
        details TEXT,
        user_name TEXT DEFAULT 'Gerente de Finanzas',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (period_id) REFERENCES periods(id)
    );
    """)
    
    conn.commit()
    seed_default_mapping_rules(conn)
    conn.close()

def seed_default_mapping_rules(conn: sqlite3.Connection):
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM mapping_rules WHERE company_id IS NULL")
    if cursor.fetchone()[0] > 0:
        return
    
    # Standard IFRS Rules
    default_rules = [
        # Activos Corrientes
        ("%caja%", "Efectivo y Equivalentes de Efectivo", "Efectivo y Equivalentes de Efectivo", "ACTIVO_CORRIENTE", "DEUDORA", "Liquidez inmediata"),
        ("%banco%", "Efectivo y Equivalentes de Efectivo", "Efectivo y Equivalentes de Efectivo", "ACTIVO_CORRIENTE", "DEUDORA", "Disponibilidades bancarias"),
        ("%fondo%fijo%", "Efectivo y Equivalentes de Efectivo", "Efectivo y Equivalentes de Efectivo", "ACTIVO_CORRIENTE", "DEUDORA", "Fondos de operación"),
        ("%deposito%plazo%", "Otros Activos Financieros Corrientes", "Otros Activos Financieros Corrientes", "ACTIVO_CORRIENTE", "DEUDORA", "Inversiones CP"),
        ("%fondo%mutuo%", "Efectivo y Equivalentes de Efectivo", "Efectivo y Equivalentes de Efectivo", "ACTIVO_CORRIENTE", "DEUDORA", "Inversiones de alta liquidez"),
        ("%cliente%", "Deudores Comerciales y Otras Cuentas por Cobrar Corrientes", "Deudores Comerciales y Otras Cuentas por Cobrar", "ACTIVO_CORRIENTE", "DEUDORA", "Créditos por ventas"),
        ("%deudores%por%venta%", "Deudores Comerciales y Otras Cuentas por Cobrar Corrientes", "Deudores Comerciales y Otras Cuentas por Cobrar", "ACTIVO_CORRIENTE", "DEUDORA", "Créditos por ventas"),
        ("%cuentas%por%cobrar%", "Deudores Comerciales y Otras Cuentas por Cobrar Corrientes", "Deudores Comerciales y Otras Cuentas por Cobrar", "ACTIVO_CORRIENTE", "DEUDORA", "Cuentas por cobrar ordinarias"),
        ("%documentos%por%cobrar%", "Deudores Comerciales y Otras Cuentas por Cobrar Corrientes", "Deudores Comerciales y Otras Cuentas por Cobrar", "ACTIVO_CORRIENTE", "DEUDORA", "Documentos mercantiles a cobrar"),
        ("%empresas%relacionadas%cobrar%", "Cuentas por Cobrar a Entidades Relacionadas Corrientes", "Cuentas por Cobrar a Entidades Relacionadas", "ACTIVO_CORRIENTE", "DEUDORA", "Partes relacionadas"),
        ("%mercaderia%", "Inventarios", "Inventarios", "ACTIVO_CORRIENTE", "DEUDORA", "Bienes de cambio"),
        ("%existencia%", "Inventarios", "Inventarios", "ACTIVO_CORRIENTE", "DEUDORA", "Bienes de cambio"),
        ("%materia%prima%", "Inventarios", "Inventarios", "ACTIVO_CORRIENTE", "DEUDORA", "Insumos y materias primas"),
        ("%productos%terminados%", "Inventarios", "Inventarios", "ACTIVO_CORRIENTE", "DEUDORA", "Inventario disponible para venta"),
        ("%iva%credito%", "Activos por Impuestos Corrientes", "Activos por Impuestos Corrientes", "ACTIVO_CORRIENTE", "DEUDORA", "Crédito fiscal tributario"),
        ("%ppm%", "Activos por Impuestos Corrientes", "Activos por Impuestos Corrientes", "ACTIVO_CORRIENTE", "DEUDORA", "Pagos provisionales mensuales"),
        ("%anticipo%proveedor%", "Otros Activos No Financieros Corrientes", "Otros Activos No Financieros Corrientes", "ACTIVO_CORRIENTE", "DEUDORA", "Anticipos comerciales"),
        ("%gasto%anticipado%", "Otros Activos No Financieros Corrientes", "Otros Activos No Financieros Corrientes", "ACTIVO_CORRIENTE", "DEUDORA", "Seguros y arriendos pagados por anticipado"),
        ("%seguros%vigentes%", "Otros Activos No Financieros Corrientes", "Otros Activos No Financieros Corrientes", "ACTIVO_CORRIENTE", "DEUDORA", "Pólizas anticipadas"),

        # Activos No Corrientes
        ("%terreno%", "Propiedades, Planta y Equipo", "Propiedades, Planta y Equipo (PP&E)", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Inmuebles y terrenos"),
        ("%edificio%", "Propiedades, Planta y Equipo", "Propiedades, Planta y Equipo (PP&E)", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Construcciones"),
        ("%maquinaria%", "Propiedades, Planta y Equipo", "Propiedades, Planta y Equipo (PP&E)", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Bienes de capital"),
        ("%muebles%utiles%", "Propiedades, Planta y Equipo", "Propiedades, Planta y Equipo (PP&E)", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Mobiliario corporativo"),
        ("%instalacion%", "Propiedades, Planta y Equipo", "Propiedades, Planta y Equipo (PP&E)", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Instalaciones de infraestructura"),
        ("%vehiculo%", "Propiedades, Planta y Equipo", "Propiedades, Planta y Equipo (PP&E)", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Flota vehicular"),
        ("%equipos%computacional%", "Propiedades, Planta y Equipo", "Propiedades, Planta y Equipo (PP&E)", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Hardware y TI"),
        ("%depreciacion%acumulada%", "Propiedades, Planta y Equipo", "Propiedades, Planta y Equipo (PP&E)", "ACTIVO_NO_CORRIENTE", "ACREEDORA", "Menor valor PP&E"),
        ("%derecho%uso%", "Activos por Derecho de Uso", "Activos por Derecho de Uso (IFRS 16)", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Arrendamientos IFRS 16"),
        ("%software%", "Activos Intangibles Distintos de la Plusvalía", "Activos Intangibles", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Licencias y software"),
        ("%marca%", "Activos Intangibles Distintos de la Plusvalía", "Activos Intangibles", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Propiedad industrial"),
        ("%impuesto%diferido%activo%", "Activos por Impuestos Diferidos", "Activos por Impuestos Diferidos", "ACTIVO_NO_CORRIENTE", "DEUDORA", "Diferencias temporarias activas"),

        # Pasivos Corrientes
        ("%proveedor%", "Cuentas Comerciales y Otras Cuentas por Pagar Corrientes", "Cuentas Comerciales y Otras Cuentas por Pagar", "PASIVO_CORRIENTE", "ACREEDORA", "Deudas operacionales"),
        ("%cuentas%por%pagar%", "Cuentas Comerciales y Otras Cuentas por Pagar Corrientes", "Cuentas Comerciales y Otras Cuentas por Pagar", "PASIVO_CORRIENTE", "ACREEDORA", "Obligaciones corrientes"),
        ("%honorarios%por%pagar%", "Cuentas Comerciales y Otras Cuentas por Pagar Corrientes", "Cuentas Comerciales y Otras Cuentas por Pagar", "PASIVO_CORRIENTE", "ACREEDORA", "Servicios profesionales"),
        ("%sobregiro%", "Otros Pasivos Financieros Corrientes", "Otros Pasivos Financieros Corrientes", "PASIVO_CORRIENTE", "ACREEDORA", "Deuda bancaria CP"),
        ("%prestamo%bancario%corto%", "Otros Pasivos Financieros Corrientes", "Otros Pasivos Financieros Corrientes", "PASIVO_CORRIENTE", "ACREEDORA", "Créditos bancarios CP"),
        ("%credito%bancario%", "Otros Pasivos Financieros Corrientes", "Otros Pasivos Financieros Corrientes", "PASIVO_CORRIENTE", "ACREEDORA", "Créditos bancarios"),
        ("%arrendamiento%corto%plazo%", "Pasivos por Arrendamiento Corrientes", "Pasivos por Arrendamiento Corrientes", "PASIVO_CORRIENTE", "ACREEDORA", "Obligaciones leasing CP"),
        ("%iva%debito%", "Pasivos por Impuestos Corrientes", "Pasivos por Impuestos Corrientes", "PASIVO_CORRIENTE", "ACREEDORA", "Débito fiscal IVA"),
        ("%retenciones%impuesto%", "Pasivos por Impuestos Corrientes", "Pasivos por Impuestos Corrientes", "PASIVO_CORRIENTE", "ACREEDORA", "Retenciones tributarias"),
        ("%impuesto%renta%por%pagar%", "Pasivos por Impuestos Corrientes", "Pasivos por Impuestos Corrientes", "PASIVO_CORRIENTE", "ACREEDORA", "Impuesto a la renta anual"),
        ("%remuneraciones%por%pagar%", "Provisiones por Beneficios a los Empleados Corrientes", "Provisiones por Beneficios a los Empleados", "PASIVO_CORRIENTE", "ACREEDORA", "Sueldos y jornales"),
        ("%leyes%sociales%por%pagar%", "Provisiones por Beneficios a los Empleados Corrientes", "Provisiones por Beneficios a los Empleados", "PASIVO_CORRIENTE", "ACREEDORA", "Seguridad social y salud"),
        ("%provision%vacaciones%", "Provisiones por Beneficios a los Empleados Corrientes", "Provisiones por Beneficios a los Empleados", "PASIVO_CORRIENTE", "ACREEDORA", "Feriado legal devengado"),
        ("%anticipo%clientes%", "Otros Pasivos No Financieros Corrientes", "Otros Pasivos No Financieros Corrientes", "PASIVO_CORRIENTE", "ACREEDORA", "Anticipos de clientes"),

        # Pasivos No Corrientes
        ("%prestamo%largo%plazo%", "Otros Pasivos Financieros No Corrientes", "Otros Pasivos Financieros No Corrientes", "PASIVO_NO_CORRIENTE", "ACREEDORA", "Deuda financiera LP"),
        ("%obligaciones%con%bancos%lp%", "Otros Pasivos Financieros No Corrientes", "Otros Pasivos Financieros No Corrientes", "PASIVO_NO_CORRIENTE", "ACREEDORA", "Deuda financiera LP"),
        ("%arrendamiento%largo%plazo%", "Pasivos por Arrendamiento No Corrientes", "Pasivos por Arrendamiento No Corrientes", "PASIVO_NO_CORRIENTE", "ACREEDORA", "Leasing financiero LP"),
        ("%impuesto%diferido%pasivo%", "Pasivos por Impuestos Diferidos", "Pasivos por Impuestos Diferidos", "PASIVO_NO_CORRIENTE", "ACREEDORA", "Diferencias temporarias pasivas"),

        # Patrimonio
        ("%capital%social%", "Capital Emitido", "Capital Emitido", "PATRIMONIO_NETO", "ACREEDORA", "Aportes de accionistas"),
        ("%capital%pagado%", "Capital Emitido", "Capital Emitido", "PATRIMONIO_NETO", "ACREEDORA", "Capital social"),
        ("%capital%", "Capital Emitido", "Capital Emitido", "PATRIMONIO_NETO", "ACREEDORA", "Capital"),
        ("%reserva%legal%", "Otras Reservas", "Otras Reservas", "PATRIMONIO_NETO", "ACREEDORA", "Reservas legales y estatutarias"),
        ("%otras%reservas%", "Otras Reservas", "Otras Reservas", "PATRIMONIO_NETO", "ACREEDORA", "Reservas patrimoniales"),
        ("%utilidad%acumulada%", "Ganancias (Pérdidas) Acumuladas", "Ganancias (Pérdidas) Acumuladas", "PATRIMONIO_NETO", "ACREEDORA", "Resultados de ejercicios anteriores"),
        ("%perdida%acumulada%", "Ganancias (Pérdidas) Acumuladas", "Ganancias (Pérdidas) Acumuladas", "PATRIMONIO_NETO", "DEUDORA", "Pérdidas de ejercicios anteriores"),
        ("%resultado%acumulado%", "Ganancias (Pérdidas) Acumuladas", "Ganancias (Pérdidas) Acumuladas", "PATRIMONIO_NETO", "ACREEDORA", "Resultados acumulados"),
        ("%resultado%ejercicio%", "Resultado del Ejercicio", "Resultado del Ejercicio", "PATRIMONIO_NETO", "ACREEDORA", "Resultado del período actual"),
        ("%utilidad%ejercicio%", "Resultado del Ejercicio", "Resultado del Ejercicio", "PATRIMONIO_NETO", "ACREEDORA", "Utilidad neta del período"),
        ("%perdida%ejercicio%", "Resultado del Ejercicio", "Resultado del Ejercicio", "PATRIMONIO_NETO", "DEUDORA", "Pérdida neta del período"),

        # Estado de Resultados / P&L
        ("%venta%", "Ingresos de Actividades Ordinarias", "Ingresos de Actividades Ordinarias", "ESTADO_RESULTADOS_INGRESOS", "ACREEDORA", "Ingresos operacionales"),
        ("%ingreso%por%servicio%", "Ingresos de Actividades Ordinarias", "Ingresos de Actividades Ordinarias", "ESTADO_RESULTADOS_INGRESOS", "ACREEDORA", "Prestación de servicios"),
        ("%costo%de%venta%", "Costo de Ventas", "Costo de Ventas", "ESTADO_RESULTADOS_COSTOS", "DEUDORA", "Costos directos de operación"),
        ("%costo%explotacion%", "Costo de Ventas", "Costo de Ventas", "ESTADO_RESULTADOS_COSTOS", "DEUDORA", "Costo de explotación"),
        ("%sueldo%", "Gastos de Administración y Ventas", "Gastos de Administración y Ventas (SG&A)", "ESTADO_RESULTADOS_GASTOS", "DEUDORA", "Gastos de personal"),
        ("%remuneracion%", "Gastos de Administración y Ventas", "Gastos de Administración y Ventas (SG&A)", "ESTADO_RESULTADOS_GASTOS", "DEUDORA", "Gastos de personal"),
        ("%arriendo%gasto%", "Gastos de Administración y Ventas", "Gastos de Administración y Ventas (SG&A)", "ESTADO_RESULTADOS_GASTOS", "DEUDORA", "Arrendamientos operacionales"),
        ("%gasto%general%", "Gastos de Administración y Ventas", "Gastos de Administración y Ventas (SG&A)", "ESTADO_RESULTADOS_GASTOS", "DEUDORA", "Gastos generales"),
        ("%servicio%basico%", "Gastos de Administración y Ventas", "Gastos de Administración y Ventas (SG&A)", "ESTADO_RESULTADOS_GASTOS", "DEUDORA", "Luz, agua, internet, telecom"),
        ("%depreciacion%ejercicio%", "Gastos de Administración y Ventas", "Gastos de Administración y Ventas (SG&A)", "ESTADO_RESULTADOS_GASTOS", "DEUDORA", "Gasto por depreciación"),
        ("%gasto%financiero%", "Costos Financieros", "Costos Financieros", "ESTADO_RESULTADOS_GASTOS", "DEUDORA", "Intereses y comisiones financieras"),
        ("%interes%pagado%", "Costos Financieros", "Costos Financieros", "ESTADO_RESULTADOS_GASTOS", "DEUDORA", "Cargas financieras"),
        ("%interes%ganado%", "Ingresos Financieros", "Ingresos Financieros", "ESTADO_RESULTADOS_INGRESOS", "ACREEDORA", "Rentas de capital e intereses a favor"),
        ("%diferencia%cambio%ganancia%", "Ingresos Financieros", "Ingresos Financieros", "ESTADO_RESULTADOS_INGRESOS", "ACREEDORA", "Fluctuaciones cambiarias favorables"),
        ("%diferencia%cambio%perdida%", "Costos Financieros", "Costos Financieros", "ESTADO_RESULTADOS_GASTOS", "DEUDORA", "Fluctuaciones cambiarias desfavorables"),
        ("%impuesto%renta%gasto%", "Gasto por Impuesto a las Ganancias", "Gasto por Impuesto a las Ganancias", "ESTADO_RESULTADOS_IMPUESTOS", "DEUDORA", "Impuesto a la renta del ejercicio")
    ]
    
    cursor.executemany("""
    INSERT INTO mapping_rules (
        account_pattern, rubro_clasificado, presentacion_auditada, categoria_ifrs, naturaleza, notes
    ) VALUES (?, ?, ?, ?, ?, ?)
    """, default_rules)
    conn.commit()

# Initialize DB when module loaded
init_db()
