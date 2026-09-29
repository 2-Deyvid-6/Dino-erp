import streamlit as st
import pandas as pd
import io
import os
import numpy as np  
import math
import re
import json
import glob
from datetime import datetime
from modulos.parser_samm import limpiar_reporte_samm
from modulos.generador_excel import generar_pdf_cliente, generar_pdf_interno_dino
from PIL import Image
from robot_api_gps import actualizar_historial_api

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Dinomontacargas - ERP Mantenimiento", 
    page_icon="🚜", 
    layout="wide"
)

# --- CARGA DE CONFIGURACIÓN ---
@st.cache_data
def cargar_configuracion():
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        st.error(f"⚠️ Error cargando config.json: {e}. Usando valores por defecto.")
        return {}

CONFIG = cargar_configuracion()

# --- FUNCIONES GLOBALES ---
def normalizar_id_universal(val):
    val_str = str(val).strip()
    match = re.search(r'\[\s*(\d+)\s*\]', val_str)
    if match: return match.group(1)
    if val_str.endswith('.0'): return val_str[:-2]
    return val_str

# --- MENÚ SUPERIOR HORIZONTAL (ENRUTADOR) ---
st.markdown("<br>", unsafe_allow_html=True)
menu_seleccionado = st.pills(
    "Navegación Principal:",
    options=[
        "📅 1. Gestión de Cronogramas",
        "🛢️ 2. Predictivo de Horómetros",
        "🚜 3. Directorio de Flota",
        "📊 4. Monitoreo de OTs por Zona",
        "📍 5. Bitácora GPS",
        "📝 6. Reportes y Descargas"
    ],
    default="📅 1. Gestión de Cronogramas",
    label_visibility="collapsed"
)

# Blindaje de seguridad: Si el usuario hace clic en el botón activo y lo desmarca, 
# forzamos a que vuelva al menú 1 en lugar de colapsar la app por valor None.
if not menu_seleccionado:
    menu_seleccionado = "📅 1. Gestión de Cronogramas"
    
st.markdown("---")
# --- BARRA LATERAL (Solo para carga de archivos) ---
with st.sidebar:
    st.header("📥 Ingreso de Datos")
    st.info("Sube el reporte de SAMM (Detalle_Visitas) para el Módulo 1.")
    archivo_samm = st.file_uploader("Cargar SAMM", type=["xls", "xlsx"], key="samm")


    
# =====================================================================
# --- CEREBRO GLOBAL: CONEXIÓN EN VIVO A SQL SERVER (SAMM) ---
# =====================================================================
if 'df_base_maestra' not in st.session_state:
    try:
        # 1. Conexión segura usando la bóveda secrets.toml apuntando a PROD
        conn = st.connection("sw_dino", type="sql")
        
        # 2. Consulta SQL: Extraemos los campos reales mapeados de view_equ_equipo
        # Incluimos intentos para traer el Subtipo Catalogo
        query_maestro = """
            SELECT TOP 1000
                equipo AS equipo,
                equipo_serial AS Serial,
                ter_tercero_tercero AS Tercero,
                ter_sucursal_sucursal AS Sucursal,
                horometroActual AS [Horometro Actual],
                [cat_catalogo.equipo_catalogo.equipo] AS Modelo,
                cat_subtipoCatalogo_subtipoCatalogo AS Subtipo_Catalogo
            FROM view_equ_equipo
            WHERE ter_tercero_tercero IS NOT NULL 
        """
        
        try:
            df_maestro = conn.query(query_maestro, ttl=3600)
        except:
             # Fallback si el nombre de la columna del Subtipo no es el estándar
             query_maestro_alt = """
                SELECT TOP 1000
                    equipo AS equipo,
                    equipo_serial AS Serial,
                    ter_tercero_tercero AS Tercero,
                    ter_sucursal_sucursal AS Sucursal,
                    horometroActual AS [Horometro Actual],
                    [cat_catalogo.equipo_catalogo.equipo] AS Modelo
                FROM view_equ_equipo
                WHERE ter_tercero_tercero IS NOT NULL 
            """
             df_maestro = conn.query(query_maestro_alt, ttl=3600)
             columnas_subtipo = [col for col in df_maestro.columns if 'subtipo' in col.lower() or 'componente' in col.lower()]
             if columnas_subtipo:
                 df_maestro['Subtipo_Catalogo'] = df_maestro[columnas_subtipo[0]]
             else:
                 df_maestro['Subtipo_Catalogo'] = ""

        # Si la columna 'Modelo' no viene en la vista, creamos un fallback
        if 'Modelo' not in df_maestro.columns:
            df_maestro['Modelo'] = "SIN REGISTRO"
            
        # 3. FILTRO DE LIMPIEZA DE COMPONENTES POR SUBTIPO
        if 'Subtipo_Catalogo' in df_maestro.columns:
            df_maestro = df_maestro[~df_maestro['Subtipo_Catalogo'].astype(str).str.contains('Componente', case=False, na=False)].copy()

        # 4. Limpiamos y preparamos los IDs para el cruce con Drive
        df_maestro['Equipo_str'] = df_maestro['equipo'].apply(normalizar_id_universal)
        equipos_conocidos = sorted(df_maestro['Equipo_str'].dropna().unique(), key=lambda x: len(str(x)), reverse=True)
        
        # 5. FUSIÓN CON EL BOT DE DRIVE
        RUTA_FICHAS = 'datos_samm/Fichas_Drive.xlsx'
        if os.path.exists(RUTA_FICHAS):
            df_fichas = pd.read_excel(RUTA_FICHAS)
            
            def extraer_numero_ficha(nombre):
                nombre_str = str(nombre).upper()
                for eq in equipos_conocidos:
                    eq_str = str(eq).upper().strip()
                    if eq_str != "" and eq_str in nombre_str: return eq 
                
                nombre_limpio = nombre_str.replace(" ", "").replace("-", "").replace("_", "")
                for eq in equipos_conocidos:
                    eq_limpio = str(eq).upper().replace(" ", "").replace("-", "").replace("_", "")
                    if eq_limpio != "" and eq_limpio in nombre_limpio: return eq
                
                match = re.search(r'(\d+)\.\w+$', str(nombre).strip(), re.IGNORECASE)
                if match: return match.group(1)
                return str(nombre)
                
            df_fichas['Equipo_str'] = df_fichas['Nombre_Archivo'].apply(extraer_numero_ficha)
            df_fichas_unicas = df_fichas.drop_duplicates(subset=['Equipo_str'], keep='first')
            
            df_maestro = pd.merge(df_maestro, df_fichas_unicas[['Equipo_str', 'Link_Ficha']], on='Equipo_str', how='left')
            
        df_maestro = df_maestro.drop(columns=['Equipo_str'], errors='ignore')
        
        st.session_state['df_base_maestra'] = df_maestro
        
    except Exception as e:
        st.error(f"⚠️ Error crítico de conexión a la Base de Datos SAMM. Detalle: {e}")
        st.stop()

# =====================================================================
# ENRUTADOR DE MÓDULOS (INICIO DE LÓGICA VISUAL)
# =====================================================================

# =====================================================================
# 📅 MÓDULO 1: GESTIÓN DE CRONOGRAMAS 
# =====================================================================
if menu_seleccionado == "📅 1. Gestión de Cronogramas":
    st.title("📅 Gestión de Cronogramas y Auditoría")
    st.markdown("---")
    
    if archivo_samm is not None:
        try:
            df_maestro_actual = st.session_state['df_base_maestra'].copy()
            df_maestro_actual['equipo_clean'] = df_maestro_actual['equipo'].apply(normalizar_id_universal)
            
            claves_maestro = df_maestro_actual['equipo_clean'].astype(str).str.strip()
            dict_tercero = dict(zip(claves_maestro, df_maestro_actual['Tercero']))
            col_ubi = 'Sucursal' if 'Sucursal' in df_maestro_actual.columns else 'sucursal' if 'sucursal' in df_maestro_actual.columns else 'Ubicacion'
            dict_sucursal = dict(zip(claves_maestro, df_maestro_actual[col_ubi].astype(str).str.strip()))

            if 'df_master' not in st.session_state or st.session_state.get('nombre_archivo') != archivo_samm.name:
                df_crudo = limpiar_reporte_samm(archivo_samm)
                df_crudo['Equipo'] = df_crudo['Equipo'].apply(normalizar_id_universal)
                
                df_crudo['Tercero_BD'] = df_crudo['Equipo'].apply(lambda x: dict_tercero.get(str(x).strip(), "SIN ASIGNAR EN BD"))
                
                if 'equipos_ignorados' not in st.session_state:
                    st.session_state['equipos_ignorados'] = set()

                def auditar_contrato_vs_bd(row):
                    equipo = str(row['Equipo']).strip() 
                    if equipo in st.session_state['equipos_ignorados']: return "VERDE"
                    
                    contrato_excel = str(row['Cliente']).upper().strip()
                    tercero_bd = str(row['Tercero_BD']).upper().strip()
                    
                    if tercero_bd in ["NAN", "SIN ASIGNAR EN BD", "NONE", ""]: return "VERDE" 
                    if contrato_excel not in tercero_bd and tercero_bd not in contrato_excel: return "ROJO" 
                    return "VERDE"
                    
                df_crudo['Alerta_Auditoria'] = df_crudo.apply(auditar_contrato_vs_bd, axis=1)

                def optimizar_fechas_por_sucursal_y_cupos(df):
                    df_opt = df.copy()
                    df_opt['Fecha_DT'] = pd.to_datetime(df_opt['Fecha_Visita'].astype(str).str.split(" ").str[0], dayfirst=True, errors='coerce')
                    df_opt['Sucursal_Maestra'] = df_opt['Equipo'].apply(lambda x: dict_sucursal.get(x, "SIN_SUCURSAL")).astype(str).str.upper().str.strip()
                    df_opt['Semana'] = df_opt['Fecha_DT'].dt.isocalendar().week
                    df_opt['Año'] = df_opt['Fecha_DT'].dt.isocalendar().year
                    cambios_realizados = 0
                    grupos = df_opt.groupby(['Cliente', 'Sucursal_Maestra', 'Año', 'Semana'])
                    
                    for (cliente, sucursal, anio, semana), grupo in grupos:
                        if pd.isna(semana) or sucursal in ["SIN_SUCURSAL", "NAN", ""]: continue
                        if len(grupo) > 1:
                            dias_disponibles = sorted(grupo['Fecha_DT'].dropna().unique())
                            if not dias_disponibles: continue
                            grupo_ordenado = grupo.sort_values('Fecha_DT')
                            idx_dia = 0
                            cupo_actual = 0
                            for idx, row in grupo_ordenado.iterrows():
                                if pd.isna(row['Fecha_DT']): continue
                                dia_asignar = dias_disponibles[idx_dia]
                                if row['Fecha_DT'] != dia_asignar:
                                    df_opt.at[idx, 'Fecha_Visita'] = dia_asignar.strftime("%d/%m/%Y")
                                    df_opt.at[idx, 'Fecha_DT'] = dia_asignar
                                    cambios_realizados += 1
                                cupo_actual += 1
                                if cupo_actual >= 4:
                                    cupo_actual = 0
                                    if idx_dia < len(dias_disponibles) - 1: idx_dia += 1

                    df_opt = df_opt.drop(columns=['Fecha_DT', 'Sucursal_Maestra', 'Semana', 'Año'])
                    return df_opt, cambios_realizados

                df_crudo, total_optimizados = optimizar_fechas_por_sucursal_y_cupos(df_crudo)
                if total_optimizados > 0: st.toast(f"🚜 Logística Activa: {total_optimizados} visitas organizadas en baldes de 4 por semana.", icon="🚜")

                st.session_state['df_master'] = df_crudo
                st.session_state['nombre_archivo'] = archivo_samm.name

            df_limpio = st.session_state['df_master']
            total_previstos = len(df_limpio)
            if total_previstos > 0:
                sin_ot = len(df_limpio[df_limpio['Color_Semantico'] == 'Rojo'])
                en_proceso = len(df_limpio[df_limpio['Color_Semantico'] == 'Amarillo'])
                finalizadas = len(df_limpio[df_limpio['Color_Semantico'] == 'Verde'])
            else:
                sin_ot, en_proceso, finalizadas = 0, 0, 0

            st.subheader("📊 Resumen General de Operaciones")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Visitas", f"{total_previstos}")
            col2.metric("🔴 Sin OT (Rojo)", f"{sin_ot}")
            col3.metric("🟡 En Proceso (Amarillo)", f"{en_proceso}")
            col4.metric("🟢 Finalizadas (Verde)", f"{finalizadas}")
            st.markdown("---")
            
            tab_buscador, tab_datos, tab_cronograma = st.tabs(["🔍 Buscador de Flota", "📋 Base de Datos General", "📅 Generar Cronogramas y Auditoría"])
            
            with tab_buscador:
                st.subheader("Radiografía por Equipo")
                lista_equipos = sorted(df_limpio['Equipo'].unique())
                equipo_buscado = st.selectbox("Selecciona el Equipo:", options=["Seleccionar..."] + lista_equipos)
                if equipo_buscado != "Seleccionar...":
                    df_equipo = df_limpio[df_limpio['Equipo'] == equipo_buscado].sort_values(by="Fecha_Visita")
                    c1, c2, c3 = st.columns(3)
                    c1.info(f"**🏢 Cliente (Contrato):**\n{df_equipo['Cliente'].iloc[0]}")
                    c2.info(f"**📍 Ubicación:**\n{df_equipo['Sucursal'].iloc[0]}")
                    c3.info(f"**🔧 Visitas:**\n{len(df_equipo)}")
                    st.dataframe(df_equipo[['Fecha_Visita', 'Mantenimiento', 'OT', 'Estado']], use_container_width=True, hide_index=True)
            
            with tab_datos:
                st.subheader("Auditoría Global de la Flota (Contrato vs Base de Datos)")
                df_anomalias = df_limpio[df_limpio['Alerta_Auditoria'] == 'ROJO'].drop_duplicates(subset=['Equipo', 'Cliente']).copy()
                
                if not df_anomalias.empty:
                    st.warning(f"🚨 Hay {len(df_anomalias)} anomalías: El Contrato programado no coincide con el Tercero real en la Base de Datos.")
                    
                    df_mostrar = df_anomalias[['Equipo', 'Cliente', 'Tercero_BD', 'Sucursal', 'Estado', 'Mantenimiento']].copy()
                    df_mostrar.rename(columns={'Cliente': 'Contrato Programado (Excel)', 'Tercero_BD': 'Tercero Real (BD)'}, inplace=True)
                    
                    st.dataframe(df_mostrar, hide_index=True, use_container_width=True)
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer: df_mostrar.to_excel(writer, sheet_name='Anomalias', index=False)
                    st.download_button("📥 Descargar Reporte Global", buffer.getvalue(), "Anomalias.xlsx", "application/vnd.ms-excel", type="primary")
                else: 
                    st.success("✅ Toda la flota programada coincide correctamente con la Base de Datos de SAMM.")
                st.write("---")
                
                st.subheader("📌 Equipos con Múltiples Visitas el Mismo Día")
                df_global_multiples = df_limpio.groupby(['Cliente', 'Equipo', 'Fecha_Visita']).size().reset_index(name='Cantidad_Visitas')
                df_global_multiples = df_global_multiples[df_global_multiples['Cantidad_Visitas'] > 1]
                if not df_global_multiples.empty:
                    st.info(f"Se detectaron {len(df_global_multiples)} casos de equipos con más de una visita programada para la misma fecha.")
                    st.dataframe(df_global_multiples, hide_index=True)
                else: st.success("✅ No hay visitas duplicadas para el mismo día en la flota.")
                st.write("---")
    
                st.subheader("🚨 Equipos Sin Cronograma de Mantenimiento")
                equipos_en_samm = df_limpio['Equipo'].unique()
                
                # 1. Filtramos sin hacer nuevas peticiones a SQL
                df_faltantes = df_maestro_actual[~df_maestro_actual['equipo_clean'].isin(equipos_en_samm)].copy()
                
                if 'Estado_Fisico' in df_faltantes.columns:
                    df_faltantes = df_faltantes[~df_faltantes['Estado_Fisico'].astype(str).str.contains('vendido', case=False, na=False)]
                
                col_mod = 'Modelo' if 'Modelo' in df_faltantes.columns else df_faltantes.columns[0]
                df_faltantes_mostrar = df_faltantes[['equipo_clean', 'Tercero', col_ubi, col_mod, 'Horometro Actual']].dropna(subset=['Tercero'])
                df_faltantes_mostrar.rename(columns={'equipo_clean': 'equipo'}, inplace=True)
                
                if not df_faltantes_mostrar.empty:
                    st.warning(f"{len(df_faltantes_mostrar)} equipos activos en BD sin visita programada.")
                    st.dataframe(df_faltantes_mostrar, hide_index=True)
                    df_word = df_faltantes_mostrar.drop(columns=['Horometro Actual'], errors='ignore')
                    
                    def generar_word_faltantes(df):
                        from docx import Document 
                        doc = Document()
                        doc.add_heading('Equipos Sin Proyección', level=1)
                        doc.add_paragraph(f'Se reportan {len(df)} equipos activos sin visita.')
                        tabla = doc.add_table(rows=1, cols=len(df.columns))
                        tabla.style = 'Table Grid'
                        hdr_cells = tabla.rows[0].cells
                        for i, col in enumerate(df.columns): hdr_cells[i].text = str(col).upper()
                        for _, fila in df.iterrows():
                            row_cells = tabla.add_row().cells
                            for i, val in enumerate(fila): row_cells[i].text = str(val)
                        buf = io.BytesIO()
                        doc.save(buf)
                        return buf.getvalue()
                        
                    st.download_button("📄 Descargar Informe (Word)", generar_word_faltantes(df_word), "Equipos_Faltantes.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
                    st.write("---")
                    
                    with st.expander("🛠️ Asignar a un Cronograma (Modo Manual Temporal)"):
                        col_eq, col_cli = st.columns(2)
                        lista_huerfanos = sorted(df_faltantes_mostrar['equipo'].unique())
                        eq_manual = col_eq.selectbox("1. Equipo sin proyección:", lista_huerfanos)
                        clientes_existentes = sorted(df_limpio['Cliente'].unique())
                        tipo_cli = col_cli.radio("2. Asignar a:", ["Cliente Existente", "Crear Nuevo Cliente"])
                        
                        if tipo_cli == "Cliente Existente": 
                            cli_destino = col_cli.selectbox("Selecciona el cliente:", clientes_existentes)
                        else: 
                            cli_destino = col_cli.text_input("Escribe el nuevo cliente:")
                            
                        num_visitas = st.number_input("¿Cuántas visitas deseas programar?", min_value=1, max_value=6, value=1)
                        cols_fechas = st.columns(num_visitas)
                        fechas_ingresadas = [cols_fechas[i].text_input(f"Visita {i+1} {'*' if i == 0 else ''}", key=f"fecha_manual_{i}", placeholder="Ej: 15/08/2026") for i in range(num_visitas)]
                        
                        if st.button("💾 Guardar y Asignar Equipo", type="primary"):
                            if not cli_destino or cli_destino.strip() == "": 
                                st.error("⚠️ El nombre del cliente es obligatorio.")
                            elif not fechas_ingresadas[0] or fechas_ingresadas[0].strip() == "": 
                                st.error("⚠️ Debes ingresar obligatoriamente la primera fecha.")
                            else:
                                info_eq = df_faltantes_mostrar[df_faltantes_mostrar['equipo'] == eq_manual].iloc[0]
                                sucursal_eq = info_eq[col_ubi] if pd.notna(info_eq[col_ubi]) else "SIN SUCURSAL"
                                tercero_bd_eq = info_eq['Tercero']
                                
                                nuevas_filas = []
                                for fecha in fechas_ingresadas:
                                    if fecha and fecha.strip() != "":
                                        nuevas_filas.append({
                                            'Equipo': eq_manual, 
                                            'Cliente': cli_destino.upper().strip(), 
                                            'Sucursal': sucursal_eq, 
                                            'Fecha_Visita': fecha.strip(), 
                                            'Estado': 'PROGRAMADO MANUAL', 
                                            'Mantenimiento': 'PREVENTIVO', 
                                            'Color_Semantico': 'Amarillo', 
                                            'Alerta_Auditoria': 'VERDE', 
                                            'Tercero_BD': tercero_bd_eq
                                        })
                                df_nuevas = pd.DataFrame(nuevas_filas)
                                st.session_state['df_master'] = pd.concat([st.session_state['df_master'], df_nuevas], ignore_index=True)
                                st.success(f"✅ ¡Equipo {eq_manual} inyectado exitosamente al contrato {cli_destino.upper()}!")
                                st.rerun()
                else: 
                    st.success("✅ Todos tienen cronograma.")
                
                col_mod = 'Modelo' if 'Modelo' in df_faltantes.columns else df_faltantes.columns[0]
                df_faltantes_mostrar = df_faltantes[['equipo_clean', 'Tercero', col_ubi, col_mod, 'Horometro Actual']].dropna(subset=['Tercero'])
                df_faltantes_mostrar.rename(columns={'equipo_clean': 'equipo'}, inplace=True)
                
                if not df_faltantes_mostrar.empty:
                    st.warning(f"{len(df_faltantes_mostrar)} equipos activos en BD sin visita programada.")
                    st.dataframe(df_faltantes_mostrar, hide_index=True)
                    df_word = df_faltantes_mostrar.drop(columns=['Horometro Actual'], errors='ignore')
                    def generar_word_faltantes(df):
                        from docx import Document 
                        doc = Document()
                        doc.add_heading('Equipos Sin Proyección', level=1)
                        doc.add_paragraph(f'Se reportan {len(df)} equipos activos sin visita.')
                        tabla = doc.add_table(rows=1, cols=len(df.columns))
                        tabla.style = 'Table Grid'
                        hdr_cells = tabla.rows[0].cells
                        for i, col in enumerate(df.columns): hdr_cells[i].text = str(col).upper()
                        for _, fila in df.iterrows():
                            row_cells = tabla.add_row().cells
                            for i, val in enumerate(fila): row_cells[i].text = str(val)
                        buf = io.BytesIO()
                        doc.save(buf)
                        return buf.getvalue()
                    st.download_button("📄 Descargar Informe (Word)", generar_word_faltantes(df_word), "Equipos_Faltantes.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
                    st.write("---")
                    with st.expander("🛠️ Asignar a un Cronograma (Modo Manual Temporal)"):
                        col_eq, col_cli = st.columns(2)
                        lista_huerfanos = sorted(df_faltantes_mostrar['equipo'].unique())
                        eq_manual = col_eq.selectbox("1. Equipo sin proyección:", lista_huerfanos)
                        clientes_existentes = sorted(df_limpio['Cliente'].unique())
                        tipo_cli = col_cli.radio("2. Asignar a:", ["Cliente Existente", "Crear Nuevo Cliente"])
                        if tipo_cli == "Cliente Existente": cli_destino = col_cli.selectbox("Selecciona el cliente:", clientes_existentes)
                        else: cli_destino = col_cli.text_input("Escribe el nuevo cliente:")
                        num_visitas = st.number_input("¿Cuántas visitas deseas programar?", min_value=1, max_value=6, value=1)
                        cols_fechas = st.columns(num_visitas)
                        fechas_ingresadas = []
                        for i in range(num_visitas):
                            f = cols_fechas[i].text_input(f"Visita {i+1} {'*' if i == 0 else ''}", key=f"fecha_manual_{i}", placeholder="Ej: 15/08/2026")
                            fechas_ingresadas.append(f)
                        if st.button("💾 Guardar y Asignar Equipo", type="primary"):
                            if not cli_destino or cli_destino.strip() == "": st.error("⚠️ El nombre del cliente es obligatorio.")
                            elif not fechas_ingresadas[0] or fechas_ingresadas[0].strip() == "": st.error("⚠️ Debes ingresar obligatoriamente la primera fecha.")
                            else:
                                info_eq = df_faltantes_mostrar[df_faltantes_mostrar['equipo'] == eq_manual].iloc[0]
                                sucursal_eq = info_eq.iloc[2] if pd.notna(info_eq.iloc[2]) else "SIN SUCURSAL"
                                nuevas_filas = []
                                for fecha in fechas_ingresadas:
                                    if fecha and fecha.strip() != "":
                                        nuevas_filas.append({'Equipo': eq_manual, 'Cliente': cli_destino.upper().strip(), 'Sucursal': sucursal_eq, 'Fecha_Visita': fecha.strip(), 'Estado': 'PROGRAMADO MANUAL', 'Mantenimiento': 'PREVENTIVO', 'Color_Semantico': 'Amarillo', 'Alerta_Auditoria': 'VERDE', 'Tercero_BD': info_eq.iloc[1]})
                                df_nuevas = pd.DataFrame(nuevas_filas)
                                st.session_state['df_master'] = pd.concat([st.session_state['df_master'], df_nuevas], ignore_index=True)
                                st.success(f"✅ ¡Equipo {eq_manual} inyectado exitosamente al contrato {cli_destino.upper()}!")
                                st.rerun()
                else: st.success("✅ Todos tienen cronograma.")

            with tab_cronograma:
                st.subheader("Auditoría y Automatización de PDF (Clientes)")
                clientes_unicos = sorted(df_limpio['Cliente'].unique())
                opciones_selector = [f"👀 {c}" if 'ROJO' in df_limpio[df_limpio['Cliente'] == c]['Alerta_Auditoria'].values else c for c in clientes_unicos]
                indice_guardado = 0
                if 'ultimo_cliente' in st.session_state:
                    for i, op in enumerate(opciones_selector):
                        if op.replace("👀 ", "") == st.session_state['ultimo_cliente']:
                            indice_guardado = i
                            break
                cliente_raw = st.selectbox("Cliente (Contrato):", opciones_selector, index=indice_guardado)
                cliente_seleccionado = cliente_raw.replace("👀 ", "")
                st.session_state['ultimo_cliente'] = cliente_seleccionado
                df_filtrado = df_limpio[df_limpio['Cliente'] == cliente_seleccionado]
                alertas_rojas = df_filtrado[df_filtrado['Alerta_Auditoria'] == 'ROJO']
                
                if not alertas_rojas.empty:
                    st.warning("👀 REVISIÓN SUGERIDA: Conflicto entre Contrato Programado y Tercero en BD.")
                    
                    df_alertas_cli = alertas_rojas[['Equipo', 'Cliente', 'Tercero_BD', 'Sucursal']].drop_duplicates().copy()
                    df_alertas_cli.rename(columns={'Cliente': 'Contrato Actual', 'Tercero_BD': 'Tercero Sugerido (BD)'}, inplace=True)
                    df_alertas_cli.insert(0, '✅ Seleccionar', False)
                    
                    df_editado_cli = st.data_editor(df_alertas_cli, hide_index=True, use_container_width=True, disabled=['Equipo', 'Contrato Actual', 'Tercero Sugerido (BD)', 'Sucursal'])
                    df_seleccionados = df_editado_cli[df_editado_cli['✅ Seleccionar'] == True]
                    
                    if not df_seleccionados.empty:
                        equipos_afectados = df_seleccionados['Equipo'].tolist()
                        def ejecutar_reasignacion(nuevo_cliente_destino):
                            mascara_mover = st.session_state['df_master']['Equipo'].isin(equipos_afectados)
                            st.session_state['df_master'].loc[mascara_mover, 'Cliente'] = nuevo_cliente_destino
                            def re_auditar(row):
                                eq_s = str(row['Equipo']).strip()
                                if eq_s in st.session_state.get('equipos_ignorados', set()): return "VERDE"
                                contrato_excel = str(row['Cliente']).upper().strip()
                                tercero_bd = str(row['Tercero_BD']).upper().strip()
                                if tercero_bd in ["NAN", "SIN ASIGNAR EN BD", "NONE", ""]: return "VERDE"
                                if contrato_excel not in tercero_bd and tercero_bd not in contrato_excel: return "ROJO"
                                return "VERDE"
                            st.session_state['df_master']['Alerta_Auditoria'] = st.session_state['df_master'].apply(re_auditar, axis=1)
                            st.rerun()

                        col1, col2, col3 = st.columns(3)
                        with col1:
                            if st.button("✅ Ignorar y Aprobar (Lista Blanca)", type="primary", use_container_width=True):
                                if 'equipos_ignorados' not in st.session_state: st.session_state['equipos_ignorados'] = set()
                                st.session_state['equipos_ignorados'].update(equipos_afectados) 
                                mascara = st.session_state['df_master']['Equipo'].isin(equipos_afectados)
                                st.session_state['df_master'].loc[mascara, 'Alerta_Auditoria'] = 'VERDE'
                                st.rerun()
                        with col2:
                            with st.expander("🔄 Reasignar a Cliente Existente"):
                                dst_existente = st.selectbox("Selecciona destino:", clientes_unicos, label_visibility="collapsed", key="sel_dst")
                                if st.button("Confirmar Reasignación", use_container_width=True):
                                    if dst_existente and dst_existente.strip() != "": ejecutar_reasignacion(dst_existente.strip().upper())
                        with col3:
                            with st.expander("📄 Crear Nuevo Contrato"):
                                dst_nuevo = st.text_input("Nombre del contrato:", label_visibility="collapsed", placeholder="Ej: TRACTOCAR CARTAGENA", key="txt_dst")
                                if st.button("Confirmar Creación", use_container_width=True):
                                    if dst_nuevo and dst_nuevo.strip() != "": ejecutar_reasignacion(dst_nuevo.strip().upper())
                st.write("---")
                
                df_visitas_multiples = df_filtrado.groupby(['Equipo', 'Fecha_Visita']).size().reset_index(name='Cantidad_Visitas')
                df_visitas_multiples = df_visitas_multiples[df_visitas_multiples['Cantidad_Visitas'] > 1]
                if not df_visitas_multiples.empty:
                    st.info("📌 **ATENCIÓN:** Los siguientes equipos tienen múltiples visitas el mismo día. (En el PDF se purgarán y mostrará 1 sola fecha).")
                    st.dataframe(df_visitas_multiples, hide_index=True)
                    
                st.write("---")
                st.subheader(f"📋 Vista Previa del Cronograma: {cliente_seleccionado}")
                df_vista_previa = df_filtrado[['Equipo', 'Sucursal', 'Fecha_Visita', 'Mantenimiento']].copy()
                df_vista_previa.insert(0, '✅ Separar', False)
                df_editado_prev = st.data_editor(df_vista_previa, hide_index=True, use_container_width=True, disabled=['Equipo', 'Sucursal', 'Fecha_Visita', 'Mantenimiento'])
                equipos_a_separar = df_editado_prev[df_editado_prev['✅ Separar'] == True]['Equipo'].tolist()
                
                if equipos_a_separar:
                    st.markdown("### ✂️ Separar Equipos Seleccionados")
                    def ejecutar_separacion(nuevo_cliente_destino):
                        mascara_sep = st.session_state['df_master']['Equipo'].isin(equipos_a_separar)
                        st.session_state['df_master'].loc[mascara_sep, 'Cliente'] = nuevo_cliente_destino
                        def re_auditar_sep(row):
                            eq_s = str(row['Equipo']).strip()
                            if eq_s in st.session_state.get('equipos_ignorados', set()): return "VERDE"
                            contrato_excel = str(row['Cliente']).upper().strip()
                            tercero_bd = str(row['Tercero_BD']).upper().strip()
                            if tercero_bd in ["NAN", "SIN ASIGNAR EN BD", "NONE", ""]: return "VERDE"
                            if contrato_excel not in tercero_bd and tercero_bd not in contrato_excel: return "ROJO"
                            return "VERDE"
                        st.session_state['df_master']['Alerta_Auditoria'] = st.session_state['df_master'].apply(re_auditar_sep, axis=1)
                        st.rerun()

                    c1_sep, c2_sep = st.columns(2)
                    with c1_sep:
                        with st.expander("🔄 Mover a Cliente Existente"):
                            dst_ex = st.selectbox("Destino:", clientes_unicos, label_visibility="collapsed", key="sel_dst_ex")
                            if st.button("Confirmar Traslado", use_container_width=True, key="btn_ex"):
                                if dst_ex and dst_ex.strip() != "": ejecutar_separacion(dst_ex.strip().upper())
                    with c2_sep:
                        with st.expander("📄 Mover a Nuevo Contrato"):
                            dst_nu = st.text_input("Nombre del contrato:", label_visibility="collapsed", placeholder="Ej: CLIENTE - SEDE NORTE", key="sel_dst_nu")
                            if st.button("Confirmar Creación", use_container_width=True, key="btn_nu"):
                                if dst_nu and dst_nu.strip() != "": ejecutar_separacion(dst_nu.strip().upper())
                                
                st.write("---")
                st.subheader("📝 Novedades y Observaciones del Mes")
                texto_novedad = st.text_area("Agrega una nota personalizada para este cliente (opcional):", placeholder="Ej: Durante este mes se observó un desgaste irregular en las llantas del equipo 408...")
                lista_combustion = []
                archivo_tracker = "datos_samm/Control_Mantenimiento.xlsx"
                if os.path.exists(archivo_tracker):
                    try:
                        df_tr = pd.read_excel(archivo_tracker)
                        col = 'EQUIPO' if 'EQUIPO' in df_tr.columns else 'Equipo'
                        lista_combustion = df_tr[col].astype(str).str.strip().tolist()
                    except: pass

                pdf_bytes = generar_pdf_cliente(df_filtrado, cliente_seleccionado, texto_novedad, lista_combustion)
                st.download_button(f"⬇️ Descargar PDF - {cliente_seleccionado}", pdf_bytes, f"Cronograma_{cliente_seleccionado}.pdf", "application/pdf", type="primary")
                st.markdown("---")
                st.subheader("🚜 Planificador Logístico Interno (Mantenimiento Dino)")
                
                palabras_excluidas = CONFIG.get("palabras_excluidas_logistica", [])
                patron_exclusion = '|'.join(palabras_excluidas) if palabras_excluidas else 'xxx_no_excluir_xxx'
                mascara_sucursal = df_limpio['Sucursal'].astype(str).str.contains(patron_exclusion, case=False, na=False)
                df_ruta_dino = df_limpio[~mascara_sucursal].copy()
                
                if not df_ruta_dino.empty:
                    total_visitas_ruta = len(df_ruta_dino)
                    horas_estimadas = total_visitas_ruta * 2
                    c1, c2 = st.columns(2)
                    c1.metric("📍 Visitas en Ruta (Excluyendo foráneos)", total_visitas_ruta)
                    c2.metric("⏱️ Horas Técnicas Estimadas (2h x Eq)", f"{horas_estimadas} hrs")
                    pdf_interno_bytes = generar_pdf_interno_dino(df_ruta_dino, horas_estimadas)
                    st.download_button(label="⚙️ Descargar Cronograma Interno (PDF)", data=pdf_interno_bytes, file_name="Cronograma_Interno_Dino.pdf", mime="application/pdf")
                    with st.expander("Ver lista de equipos incluidos en esta ruta"):
                        st.dataframe(df_ruta_dino[['Cliente', 'Equipo', 'Sucursal', 'Fecha_Visita']].sort_values(by=['Sucursal', 'Fecha_Visita']), hide_index=True)
                else: st.warning("No se encontraron equipos para esta ruta después de aplicar las exclusiones.")

        except Exception as e: st.error(f"Error crítico al procesar el archivo. Detalles: {e}")
    else: st.info("👈 Sube el reporte de SAMM para gestionar los cronogramas.")

# =====================================================================
# 🟦 MÓDULO 2: PREDICTIVO DE HORÓMETROS (Aceites y Filtros)
# =====================================================================
elif menu_seleccionado == "🛢️ 2. Predictivo de Horómetros":
    st.title("🛢️ Control Predictivo de Mantenimientos")
    st.markdown("---")
    
    df_maestro = st.session_state['df_base_maestra'].copy()
    
    def limpiar_horometro_base(val):
        if pd.isna(val): return None
        val_str = str(val).strip().replace(',', '.')
        try:
            num = float(val_str)
            while num > 40000: num = num / 10.0
            return num
        except: return None
            
    df_maestro['Horometro Actual'] = df_maestro['Horometro Actual'].apply(limpiar_horometro_base)
    df_maestro = df_maestro.dropna(subset=['Horometro Actual'])
    df_maestro['equipo_str'] = df_maestro['equipo'].apply(normalizar_id_universal)
    
    archivo_tracker = "datos_samm/Control_Mantenimiento.xlsx"
    if os.path.exists(archivo_tracker):
        df_tracker_crudo = pd.read_excel(archivo_tracker)
        df_tracker = df_tracker_crudo.copy()
        if 'EQUIPO' in df_tracker.columns: df_tracker.rename(columns={'EQUIPO': 'Equipo'}, inplace=True)
        if 'Horometro de cambio deaceite' in df_tracker.columns: df_tracker.rename(columns={'Horometro de cambio deaceite': 'Ultimo_Mantenimiento'}, inplace=True)
        if 'ESTADO INSUMOS' not in df_tracker.columns: df_tracker['Estado_Insumos'] = "Al día"
        else:
            df_tracker.rename(columns={'ESTADO INSUMOS': 'Estado_Insumos'}, inplace=True)
            df_tracker['Estado_Insumos'] = df_tracker['Estado_Insumos'].fillna("Al día")
        if 'Horometro_Base_Ciclo' not in df_tracker.columns: df_tracker['Horometro_Base_Ciclo'] = df_tracker['Ultimo_Mantenimiento']
        df_tracker['Equipo'] = df_tracker['Equipo'].apply(normalizar_id_universal)
        df_tracker = df_tracker[~df_tracker['Equipo'].str.upper().isin(['NAN', 'NONE', 'NA', ''])]
        df_maestro = df_maestro[~df_maestro['equipo_str'].str.upper().isin(['NAN', 'NONE', 'NA', ''])]
    else:
        st.warning("⚠️ No se encontró el archivo 'Control_Mantenimiento.xlsx' en 'datos_samm'.")
        st.stop()
        
    col_ubi_pred = 'Sucursal' if 'Sucursal' in df_maestro.columns else 'sucursal' if 'sucursal' in df_maestro.columns else 'Ubicacion'
    df_predictivo = pd.merge(df_tracker, df_maestro[['equipo_str', 'Tercero', col_ubi_pred, 'Modelo', 'Horometro Actual']], left_on='Equipo', right_on='equipo_str', how='inner')
    df_predictivo['Ultimo_Mantenimiento'] = pd.to_numeric(df_predictivo['Ultimo_Mantenimiento'], errors='coerce').fillna(0)
    df_predictivo['Horometro_Base_Ciclo'] = pd.to_numeric(df_predictivo['Horometro_Base_Ciclo'], errors='coerce').fillna(df_predictivo['Ultimo_Mantenimiento'])
    
    def corregir_coma_flotante(row):
        actual_original = row['Horometro Actual']
        ultimo = row['Ultimo_Mantenimiento']
        if pd.isna(actual_original) or pd.isna(ultimo) or ultimo == 0: return actual_original
        actual = actual_original
        if (actual - ultimo) > 1000:
            for _ in range(2): 
                prueba = actual / 10.0
                if prueba >= (ultimo - 500) and (prueba - ultimo) < 1500: return prueba 
                actual = prueba
        return actual_original
        
    df_predictivo['Horometro Actual'] = df_predictivo.apply(corregir_coma_flotante, axis=1)
    df_predictivo['Proximo_Mantenimiento'] = df_predictivo['Ultimo_Mantenimiento'] + 250
    df_predictivo['Horas_Faltantes'] = df_predictivo['Proximo_Mantenimiento'] - df_predictivo['Horometro Actual']
    
    def calcular_nivel(row):
        faltan = row['Horas_Faltantes']
        if faltan < -50: return "NIVEL MÁXIMO (Requiere Cambio TOTAL por Atraso)"
        horas_acumuladas = row['Proximo_Mantenimiento'] - row['Horometro_Base_Ciclo']
        if horas_acumuladas <= 0: return "NIVEL 1 (Filtros Básicos Motor)"
        ciclo_exacto = round(horas_acumuladas / 250) * 250
        if ciclo_exacto % 2000 == 0: return "NIVEL 4 (Básico + Diferencial + Hidráulico)"
        if ciclo_exacto % 1000 == 0: return "NIVEL 3 (Básico + Caja Int + Correa)"
        if ciclo_exacto % 500 == 0: return "NIVEL 2 (Básico + Caja Ext + Frenos)"
        return "NIVEL 1 (Filtros Básicos Motor)"
        
    df_predictivo['Tipo_Mantenimiento'] = df_predictivo.apply(calcular_nivel, axis=1)
    
    def semaforo(faltan):
        if faltan <= 0: return "🔴 VENCIDO (Urgente)"
        if faltan <= 30: return "🟡 PREVENTIVO (Pedir Insumos)"
        return "🟢 ÓPTIMO"
        
    df_predictivo['Alerta'] = df_predictivo['Horas_Faltantes'].apply(semaforo)
    
    # --- CORRECCIÓN INTELIGENTE DE ESTADO DE INSUMOS ---
    def corregir_estado_insumos(row):
        estado_actual = str(row.get('Estado_Insumos', 'Al día')).strip()
        alerta_actual = str(row.get('Alerta', ''))
        
        # Si el Excel dice "Al día" o está vacío, pero la matemática dice que está vencido/preventivo:
        if estado_actual.lower() in ["al día", "al dia", "nan", ""]:
            if "🔴" in alerta_actual:
                return "❌ Sin gestionar (Urgente)"
            elif "🟡" in alerta_actual:
                return "⚠️ Requiere Solicitud"
        
        return estado_actual

    df_predictivo['Estado_Insumos'] = df_predictivo.apply(corregir_estado_insumos, axis=1)
    # ---------------------------------------------------
    
    st.subheader("🚨 Panel de Alertas y Trámites")
    df_alertas = df_predictivo[df_predictivo['Alerta'].str.contains('🔴|🟡')].sort_values(by='Horas_Faltantes')
    
    if not df_alertas.empty:
        columnas_ver = ['Equipo', 'Tercero', 'Horometro Actual', 'Horas_Faltantes', 'Tipo_Mantenimiento', 'Alerta', 'Estado_Insumos']
        st.dataframe(df_alertas[columnas_ver], hide_index=True)
        buffer_alertas = io.BytesIO()
        with pd.ExcelWriter(buffer_alertas, engine='xlsxwriter') as writer: df_alertas[columnas_ver].to_excel(writer, sheet_name='Solicitud_FiltroEGMs', index=False)
        st.download_button(label="📥 Descargar Reporte para Compras (Excel)", data=buffer_alertas.getvalue(), file_name="Solicitud_Insumos_Mantenimiento.xlsx", mime="application/vnd.ms-excel")
        st.markdown("---")
        st.subheader("⚙️ Gestión de Estado del Mantenimiento")
        c1, c2, c3 = st.columns([2, 2, 1])
        with c1: equipo_revisado = st.selectbox("1. Selecciona el equipo:", df_alertas['Equipo'].tolist())
        with c2: nuevo_estado = st.selectbox("2. Actualizar estado a:", ["Solicitado", "Pendiente por instalar", "✅ Cambio BÁSICO Realizado", "🚨 Cambio TOTAL Realizado (Reseteo)"])
        with c3:
            st.write("")
            st.write("")
            if st.button("Guardar Estado", type="primary", use_container_width=True):
                if "Realizado" in nuevo_estado:
                    nuevo_horometro = df_predictivo.loc[df_predictivo['Equipo'] == equipo_revisado, 'Horometro Actual'].values[0]
                    df_tracker.loc[df_tracker['Equipo'] == equipo_revisado, 'Ultimo_Mantenimiento'] = nuevo_horometro
                    df_tracker.loc[df_tracker['Equipo'] == equipo_revisado, 'Estado_Insumos'] = "Al día"
                    if "TOTAL" in nuevo_estado:
                        df_tracker.loc[df_tracker['Equipo'] == equipo_revisado, 'Horometro_Base_Ciclo'] = nuevo_horometro
                        st.success(f"¡Reseteo exitoso! El equipo {equipo_revisado} inicia un nuevo ciclo.")
                    else: st.success(f"Mantenimiento básico registrado.")
                    from datetime import datetime
                    if 'Fecha de cambio aceite' in df_tracker.columns: df_tracker.loc[df_tracker['Equipo'] == equipo_revisado, 'Fecha de cambio aceite'] = datetime.now().strftime("%Y-%m-%d")
                else:
                    df_tracker.loc[df_tracker['Equipo'] == equipo_revisado, 'Estado_Insumos'] = nuevo_estado
                    st.success(f"Estado actualizado a: {nuevo_estado}")
                df_salida = df_tracker.rename(columns={'Equipo': 'EQUIPO', 'Ultimo_Mantenimiento': 'Horometro de cambio deaceite', 'Estado_Insumos': 'ESTADO INSUMOS'})
                
                # --- BLINDAJE ATÓMICO ---
                # Escribimos en un archivo temporal para no bloquear a quien esté leyendo
                archivo_temp = archivo_tracker.replace(".xlsx", "_temp.xlsx")
                df_salida.to_excel(archivo_temp, index=False)
                # Reemplazo a nivel sistema operativo (Milisegundos, 100% seguro)
                os.replace(archivo_temp, archivo_tracker)
                
                st.rerun()
                df_salida.to_excel(archivo_tracker, index=False)
                st.rerun()
    else: st.success("✅ Toda la flota está en estado óptimo.")
    with st.expander("📊 Ver Flota Completa (Combustión) y Estados"):
        st.dataframe(df_predictivo[['Equipo', 'Tercero', 'Horometro Actual', 'Ultimo_Mantenimiento', 'Horometro_Base_Ciclo', 'Horas_Faltantes', 'Tipo_Mantenimiento', 'Estado_Insumos']], hide_index=True)

# =====================================================================
# 🚜 MÓDULO 3: DIRECTORIO DE FLOTA Y NOVEDADES (DATOS REALES SAMM)
# =====================================================================
elif menu_seleccionado == "🚜 3. Directorio de Flota":
    st.title("🚜 Directorio Global de Flota y Novedades")
    
    df_dir = st.session_state['df_base_maestra'].copy()
    col_ubi_dir = 'Sucursal' if 'Sucursal' in df_dir.columns else 'sucursal' if 'sucursal' in df_dir.columns else 'Ubicacion'

    # --- 1. CONSULTAS REALES DESDE SQL SERVER ---
    try:
        conn = st.connection("sw_dino", type="sql")
        
        # A) Extraer estado físico para aislar Vendidos y Fuera de Servicio
        try:
            query_estados = "SELECT equipo, estadoEquipo AS Estado_Fisico FROM view_equ_equipo"
            df_estados = conn.query(query_estados, ttl=30)
            mapa_estados_fisicos = dict(zip(df_estados['equipo'].astype(str).str.strip(), df_estados['Estado_Fisico'].astype(str).str.strip()))
        except:
            try:
                query_estados_alt = "SELECT equipo, equ_estadoEquipo_estadoEquipo AS Estado_Fisico FROM view_equ_equipo"
                df_estados = conn.query(query_estados_alt, ttl=30)
                mapa_estados_fisicos = dict(zip(df_estados['equipo'].astype(str).str.strip(), df_estados['Estado_Fisico'].astype(str).str.strip()))
            except:
                mapa_estados_fisicos = {}

        # B) Extraer el historial de Novedades / Solicitudes
        query_todas_solicitudes = """
            SELECT 
                equ_equipo_equipo AS Equipo,
                fechaCreacion AS Fecha,
                [documento.solicitud] AS Novedad,
                solicitante AS Solicitante,
                doc_documento_solicitud_doc_estadoTipoDocumento_estadoTipoDocumento AS Estado
            FROM view_doc_documento_solicitud
            ORDER BY fechaCreacion DESC
        """
        df_solicitudes_raw = conn.query(query_todas_solicitudes, ttl=30)
        
        if not df_solicitudes_raw.empty:
            df_solicitudes_raw['Fecha'] = pd.to_datetime(df_solicitudes_raw['Fecha']).dt.strftime('%d/%m/%Y %I:%M %p')
            
            def clasificar_estado(est_str):
                est_str = str(est_str).lower().strip()
                if any(k in est_str for k in ['nueva', 'solicitado', 'abierto', 'pendiente', 'registrado']):
                    return "🔴 Novedad Nueva"
                elif any(k in est_str for k in ['proceso', 'asignado', 'en ejecucion', 'programado']):
                    return "🟡 En Proceso"
                else:
                    return "🟢 Sin Novedad"
            
            # 1. Clasificamos TODOS los estados primero
            df_solicitudes_raw['Categoria_Estado_Temp'] = df_solicitudes_raw['Estado'].apply(clasificar_estado)
            
            # 2. Tomamos solo los 3 registros más recientes por cada equipo
            df_top3 = df_solicitudes_raw.groupby('Equipo').head(3)
            
            # 3. Regla de negocio: La novedad "Nueva" tiene jerarquía absoluta sobre "En Proceso"
            def consolidar_prioridad(estados):
                lista = estados.tolist()
                if "🔴 Novedad Nueva" in lista:
                    return "🔴 Novedad Nueva"
                elif "🟡 En Proceso" in lista:
                    return "🟡 En Proceso"
                return "🟢 Sin Novedad"
            
            # 4. Aplicamos la regla y generamos el mapa final
            df_final_estados = df_top3.groupby('Equipo')['Categoria_Estado_Temp'].apply(consolidar_prioridad).reset_index()
            mapa_estados = dict(zip(df_final_estados['Equipo'].astype(str).str.strip(), df_final_estados['Categoria_Estado_Temp']))
        else:
            mapa_estados = {}
            df_solicitudes_raw = pd.DataFrame()

    except Exception as e:
        st.error(f"⚠️ Error conectando a la base de datos de SAMM: {e}")
        mapa_estados = {}
        mapa_estados_fisicos = {}
        df_solicitudes_raw = pd.DataFrame()


    # --- 2. LÓGICA MATEMÁTICA Y DESCARTES ---
    df_dir['Estado_Fisico'] = df_dir['equipo'].astype(str).str.strip().map(mapa_estados_fisicos).fillna("Activo")
    
    # 1. EXPULSAR VENDIDOS
    df_dir = df_dir[~df_dir['Estado_Fisico'].str.contains('vendido', case=False, na=False)]
    
    # 2. Asignar novedades a los que quedan
    df_dir['Categoria_Estado'] = df_dir['equipo'].astype(str).str.strip().map(mapa_estados).fillna("🟢 Sin Novedad")

    # 3. AISLAR Y RESTAR "FUERA DE SERVICIO" (Respetando la jerarquía de novedades)
    mascara_fuera_servicio = df_dir['Estado_Fisico'].str.contains('fuera de servicio', case=False, na=False)
    
    # Solo cambiamos la visualización a "⚪ Fuera de Servicio" si el equipo NO tiene un ticket abierto.
    # Si tiene una Novedad, mantenemos la Novedad para que aparezca en el radar del técnico.
    df_dir.loc[mascara_fuera_servicio & (df_dir['Categoria_Estado'] == '🟢 Sin Novedad'), 'Categoria_Estado'] = "⚪ Fuera de Servicio"


    # --- 3. INDICADORES KPI ---
    total_equipos = len(df_dir)
    con_novedad = len(df_dir[df_dir['Categoria_Estado'] == '🔴 Novedad Nueva'])
    
    # El KPI lo contamos directamente de la variable física, asegurando que el número no cambie 
    # aunque el equipo se muestre visualmente como 'Novedad Nueva'.
    fuera_servicio = len(df_dir[mascara_fuera_servicio])
    
    en_proceso = len(df_dir[df_dir['Categoria_Estado'] == '🟡 En Proceso'])
    sanos = len(df_dir[df_dir['Categoria_Estado'] == '🟢 Sin Novedad'])
    atendidas_total = en_proceso + sanos
    atendidas_total = en_proceso + sanos
    
    pct_novedad = (con_novedad / total_equipos * 100) if total_equipos > 0 else 0
    pct_atendidas = (atendidas_total / total_equipos * 100) if total_equipos > 0 else 0
    pct_fuera = (fuera_servicio / total_equipos * 100) if total_equipos > 0 else 0
    
    c_kpi1, c_kpi2, c_kpi3, c_kpi4 = st.columns(4)
    
    c_kpi1.metric("Fuera de Servicio", f"⚪ {fuera_servicio} equipos", f"{pct_fuera:.1f}% inactiva", delta_color="off")
    c_kpi2.metric("Novedades Nuevas", f"🔴 {con_novedad} equipos", f"-{pct_novedad:.1f}% de la flota", delta_color="normal")
    c_kpi3.metric("Atendidas / Sin Novedad", f"🟢 {atendidas_total} equipos", f"+{pct_atendidas:.1f}% operativa", delta_color="normal")
    c_kpi4.metric("Base Maestra", f"🚜 {total_equipos} equipos", "Total Flota Activa", delta_color="off")
    
    st.markdown("---")

    # --- 4. BUSCADOR SEGMENTADO ---
    st.subheader("🔍 Buscador Segmentado")
    
    c1, c2, c3 = st.columns(3)
    with c1:
        lista_clientes = ["Todos"] + sorted(df_dir['Tercero'].dropna().astype(str).unique())
        filtro_cliente = st.selectbox("Filtrar por Cliente:", lista_clientes, key="cli_global")
    with c2:
        lista_modelos_global = ["Todos"] + sorted(df_dir['Modelo'].dropna().astype(str).unique())
        filtro_modelo_global = st.selectbox("Filtrar por Modelo (Global):", lista_modelos_global, key="mod_global")
    with c3:
        filtro_equipo = st.text_input("Equipo o Serial:", key="eq_global")

    filtro_sucursal = "Todas"
    filtro_modelo_cliente = "Todos"
    
    if filtro_cliente != "Todos":
        cs1, cs2, cs3, cs4 = st.columns(4)
        with cs1:
            sucursales_cliente = df_dir[df_dir['Tercero'].astype(str) == filtro_cliente][col_ubi_dir]
            lista_sucursales = ["Todas"] + sorted(sucursales_cliente.dropna().astype(str).unique())
            filtro_sucursal = st.selectbox("↳ Sucursal del Cliente:", lista_sucursales, key="suc_cliente")
        with cs2:
            df_temp_modelos = df_dir[df_dir['Tercero'].astype(str) == filtro_cliente]
            if filtro_sucursal != "Todas":
                df_temp_modelos = df_temp_modelos[df_temp_modelos[col_ubi_dir].astype(str) == filtro_sucursal]
            lista_modelos_cli = ["Todos"] + sorted(df_temp_modelos['Modelo'].dropna().astype(str).unique())
            filtro_modelo_cliente = st.selectbox("↳ Modelo del Cliente:", lista_modelos_cli, key="mod_cliente")

    st.markdown("---")
    st.write("📌 **Filtro Estricto y Ordenamiento**")
    cf1, cf2 = st.columns(2)
    with cf1:
        estados_disponibles = ["Todos", "🔴 Novedad Nueva", "🟡 En Proceso", "🟢 Sin Novedad", "⚪ Fuera de Servicio"]
        filtro_estricto_estado = st.selectbox("1. Filtrar (Eliminar lo no seleccionado):", estados_disponibles, key="filtro_estado")
    with cf2:
        opciones_orden = [
            "Por número, menor", 
            "Por número, mayor",
            "Estado: Nuevas",
            "Estado: En proceso",
            "Estado: Sin novedad",
            "Estado: Fuera de servicio"
        ]
        orden_seleccionado = st.selectbox("2. Organizar lista por:", opciones_orden, key="orden_estado")

    # ==========================================
    # APLICACIÓN DE FILTROS EN CASCADA
    # ==========================================
    df_filtrado = df_dir.copy()
    
    # Filtros base
    if filtro_cliente != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Tercero'].astype(str) == filtro_cliente]
    if filtro_sucursal != "Todas":
        df_filtrado = df_filtrado[df_filtrado[col_ubi_dir].astype(str) == filtro_sucursal]
    if filtro_cliente != "Todos":
        if filtro_modelo_cliente != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Modelo'].astype(str) == filtro_modelo_cliente]
    else:
        if filtro_modelo_global != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Modelo'].astype(str) == filtro_modelo_global]
    if filtro_equipo:
        df_filtrado = df_filtrado[df_filtrado['equipo'].astype(str).str.contains(filtro_equipo, case=False)]

    # 1. FILTRADO ESTRICTO POR ESTADO
    if filtro_estricto_estado != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Categoria_Estado'] == filtro_estricto_estado]

    # 2. LÓGICA DE ORDENAMIENTO DE LA LISTA
    if orden_seleccionado == "Por número, menor":
        df_filtrado = df_filtrado.sort_values(by='equipo', ascending=True)
    elif orden_seleccionado == "Por número, mayor":
        df_filtrado = df_filtrado.sort_values(by='equipo', ascending=False)
    else:
        estado_prioridad = None
        if "Nuevas" in orden_seleccionado: estado_prioridad = "🔴 Novedad Nueva"
        elif "En proceso" in orden_seleccionado: estado_prioridad = "🟡 En Proceso"
        elif "Sin novedad" in orden_seleccionado: estado_prioridad = "🟢 Sin Novedad"
        elif "Fuera de servicio" in orden_seleccionado: estado_prioridad = "⚪ Fuera de Servicio"
        
        if estado_prioridad:
            # Crea una columna temporal de prioridad. Los que coincidan son True (1), el resto False (0).
            df_filtrado['es_prioridad'] = df_filtrado['Categoria_Estado'] == estado_prioridad
            # Ordena primero por prioridad (descendente para que True quede arriba) y luego por equipo
            df_filtrado = df_filtrado.sort_values(by=['es_prioridad', 'equipo'], ascending=[False, True])
            df_filtrado = df_filtrado.drop(columns=['es_prioridad'])

    st.markdown("---")
    
    # --- 5. LISTA DE EQUIPOS TIPO "CORTINA" (NUEVO DISEÑO CON FOTO) ---
    if df_filtrado.empty:
        st.warning("No se encontraron equipos con los filtros aplicados.")
    else:
        st.markdown(f"### Resultados de Búsqueda ({len(df_filtrado)} equipos)")
        
        # 📂 Rutas Base para los archivos locales actualizadas y estandarizadas
        ruta_fotos = r"datos_samm\fotos_equipos" 
        
        for idx, row in df_filtrado.iterrows():
            equipo = str(row['equipo']).strip()
            tercero = str(row.get('Tercero', 'Sin Cliente'))
            sucursal = str(row.get(col_ubi_dir, 'Sin Sucursal'))
            modelo = str(row.get('Modelo', 'Sin Modelo'))
            estado = row.get('Categoria_Estado', '🟢 Sin Novedad')
            
            posibles_nombres = [c for c in df_filtrado.columns if 'serial' in str(c).lower() or 'serie' in str(c).lower() or 'chasis' in str(c).lower()]
            columna_serial_real = posibles_nombres[0] if posibles_nombres else None
            serial = row.get(columna_serial_real, 'No registrado') if columna_serial_real else 'No encontrado'
            
            horometro = row.get('Horometro Actual', 'No registrado')
            horometro_txt = horometro if pd.notna(horometro) else 'No registrado'
            
            link_ficha = row.get('Link_Ficha', '')
            
            # Archivos esperados (Blindaje para mayúsculas y minúsculas)
            archivo_foto_jpg = os.path.join(ruta_fotos, f"{equipo}.jpg")
            archivo_foto_jpg_mayus = os.path.join(ruta_fotos, f"{equipo}.JPG")
            archivo_foto_png = os.path.join(ruta_fotos, f"{equipo}.png")
            
            # TÍTULO DEL EXPANDER
            titulo_cortina = f"Equipo {equipo} {'&nbsp;'*40} | {'&nbsp;'*40} {estado}"
            
            with st.expander(titulo_cortina):
                
                # 📐 ESTRUCTURA DE 2 COLUMNAS (Info + Foto)
                col_info, col_foto = st.columns([3, 2])
                
                # 👉 COLUMNA IZQUIERDA: Info y Botones
                with col_info:
                    st.write(f"🏢 **Cliente:** {tercero}")
                    st.write(f"📍 **Ubicación:** {sucursal}")
                    st.write(f"🏷️ **Modelo:** {modelo}")
                    st.write(f"🔢 **Serial:** {serial}")
                    st.write(f"⏱️ **Horómetro:** {horometro_txt}")
                    
                    st.write("")
                    st.markdown("**📚 Documentación Técnica**")
                    
                    # Ficha Técnica (Con llave 'key' para evitar el error de duplicados si está inactivo)
                    if pd.notna(link_ficha) and str(link_ficha).strip() != "":
                        st.link_button("📄 Ver Ficha (Drive)", str(link_ficha).strip(), type="primary", use_container_width=True)
                    else:
                        st.button("📄 Ficha (Sin Drive)", disabled=True, use_container_width=True, key=f"no_ficha_{equipo}")

                # 👉 COLUMNA DERECHA: Fotografía Inteligente
                with col_foto:
                    # Determinamos cuál de las 3 rutas es la que existe
                    ruta_img = None
                    if os.path.exists(archivo_foto_jpg): ruta_img = archivo_foto_jpg
                    elif os.path.exists(archivo_foto_jpg_mayus): ruta_img = archivo_foto_jpg_mayus
                    elif os.path.exists(archivo_foto_png): ruta_img = archivo_foto_png

                    if ruta_img:
                        try:
                            # Abrimos la imagen para inspeccionar sus medidas reales
                            img = Image.open(ruta_img)
                            ancho, alto = img.size
                            
                            if alto > ancho:
                                # 📱 Es VERTICAL: Le ponemos un ancho fijo para que no rompa la pantalla
                                st.image(img, caption=f"Montacargas #{equipo}", width=250)
                            else:
                                # 🖥️ Es HORIZONTAL o CUADRADA: Dejamos que llene la columna
                                st.image(img, caption=f"Montacargas #{equipo}", use_container_width=True)
                        except Exception as e:
                            # Fallback de seguridad en caso de que la imagen esté corrupta
                            st.image(ruta_img, caption=f"Montacargas #{equipo}", width=250)
                    else:
                        st.info("📷 Sin fotografía disponible localmente")
                
                # 🛠️ HISTORIAL SAMM (Se mantiene debajo de la foto)
                st.write("**Historial de Novedades y Solicitudes (SAMM)**")
                if not df_solicitudes_raw.empty:
                    df_nov_eq = df_solicitudes_raw[df_solicitudes_raw['Equipo'].astype(str).str.strip() == equipo].drop(columns=['Equipo'])
                    
                    if not df_nov_eq.empty:
                        st.dataframe(df_nov_eq, use_container_width=True, hide_index=True)
                    else:
                        st.info("🟢 Sin solicitudes ni novedades registradas para este equipo.")
                else:
                    st.info("🟢 Sin solicitudes ni novedades registradas para este equipo.")

# =====================================================================
# 📊 MÓDULO 4: MONITOREO DE OTs (ZONA Y TÉCNICO MULTI-ASIGNACIÓN)
# =====================================================================
elif menu_seleccionado == "📊 4. Monitoreo de OTs por Zona":
    st.title("📊 Monitoreo General de OTs (Zonas y Técnicos)")
    st.markdown("---")
    
    MAPA_TECNICOS = CONFIG.get("mapa_tecnicos", {})
    
    try:
        conn = st.connection("sw_dino", type="sql")
        
        # 1️⃣ EXTRACCIÓN CON LEFT JOIN (Columnas Equipo y Sucursal con nombres reales)
        query_ots = """
            SELECT 
                ot.doc_documento_ot_documento_numero AS OT_Num,
                ot.doc_documento_ot_doc_estadoTipoDocumento_estadoTipoDocumento AS Estado,
                ot.gen_zona_zona AS Zona, 
                ot.doc_documento_ot_ter_tercero_cliente_tercero AS Cliente,
                ot.ter_sucursal_sucursal AS Sucursal,
                ot.equ_equipo_equipo AS Equipo,
                prog.seg_usuario_usuario AS Tecnico_ID
            FROM view_doc_documento_ot AS ot
            LEFT JOIN view_ort_programacion AS prog
                ON ot.id = prog.[id_documento.ot]
               AND (prog.id_motivoCancelacion = 0 OR prog.id_motivoCancelacion IS NULL)
            WHERE ot.doc_documento_ot_doc_estadoTipoDocumento_estadoTipoDocumento IN ('Programada', 'Aprobada')
              AND ot.doc_documento_ot_doc_subtipoDocumento_subtipoDocumento = 'Orden de trabajo'
        """
        
        df_ots = conn.query(query_ots, ttl=60)
        
        if not df_ots.empty:
            
            # Limpieza básica de datos
            df_ots['Zona'] = df_ots['Zona'].fillna('SIN ZONA').astype(str).str.upper().str.strip()
            df_ots['Cliente'] = df_ots['Cliente'].fillna('SIN CLIENTE').astype(str).str.upper().str.strip()
            df_ots['Tecnico_ID_Clean'] = df_ots['Tecnico_ID'].fillna('SIN TÉCNICO ASIGNADO').astype(str).str.upper().str.strip()
            df_ots['Tecnico'] = df_ots['Tecnico_ID_Clean'].map(MAPA_TECNICOS).fillna(df_ots['Tecnico_ID_Clean'])
            
            # Universos
            df_prog = df_ots[df_ots['Estado'] == 'Programada'].copy()
            df_apro = df_ots[df_ots['Estado'] == 'Aprobada'].copy()
            
            total_prog_unicas = df_prog['OT_Num'].nunique()
            total_apro_unicas = df_apro['OT_Num'].nunique()
            
            c_met1, c_met2 = st.columns(2)
            c_met1.metric("Total OTs Programadas", f"🔴 {total_prog_unicas}", delta="Rezago Operativo", delta_color="inverse")
            c_met2.metric("Total OTs Aprobadas", f"⚠️ {total_apro_unicas}", delta="Proceso incompleto", delta_color="off")
            st.markdown("---")
            
            # ---> BLOQUE A: POR ZONA <---
            st.subheader("📍 Desglose Logístico por Zona")
            col_izq_z, col_der_z = st.columns(2)
            
            def renderizar_zonas(df_universo, titulo, icono):
                if df_universo.empty: return
                df_zona_unicas = df_universo.drop_duplicates(subset=['OT_Num'])
                total_unicos_zona = len(df_zona_unicas)
                
                zonas = df_zona_unicas.groupby('Zona').size().reset_index(name='Total').sort_values('Total', ascending=False)
                for _, z in zonas.iterrows():
                    pct = (z['Total'] / total_unicos_zona) * 100
                    with st.expander(f"{icono} {z['Zona']} — {z['Total']} OTs ({pct:.1f}%)"):
                        cli_zona = df_zona_unicas[df_zona_unicas['Zona'] == z['Zona']].groupby('Cliente').size().reset_index(name='T_Cli').sort_values('T_Cli', ascending=False)
                        for _, c in cli_zona.iterrows():
                            pct_c = (c['T_Cli'] / z['Total']) * 100
                            st.write(f"🏢 **{c['Cliente']}** ({pct_c:.1f}% / {c['T_Cli']} OTs)")
                            st.progress(float(pct_c) / 100.0)

            with col_izq_z:
                st.markdown("**🔴 Programadas**")
                renderizar_zonas(df_prog, "Programadas", "📍")
            with col_der_z:
                st.markdown("**⚠️ Aprobadas**")
                renderizar_zonas(df_apro, "Aprobadas", "🟡")
                
            st.markdown("---")
            
            ## ---> BLOQUE B: POR TÉCNICO (Multi-Asignación) <---
            st.subheader("👨‍🔧 Desglose Operativo por Técnico (Solo Programadas)")
            st.markdown(f"**🔍 Total de OTs físicas:** `{total_prog_unicas} OTs` *(Las OTs compartidas sumarán en la agenda de cada técnico asignado)*")
            
            col_tec, _ = st.columns([1, 0.01])
            
            with col_tec:
                if not df_prog.empty:
                    df_prog_tec = df_prog.drop_duplicates(subset=['OT_Num', 'Tecnico'])
                    tecnicos = df_prog_tec.groupby('Tecnico').size().reset_index(name='Total_Tec').sort_values('Total_Tec', ascending=False)
                    
                    for _, t in tecnicos.iterrows():
                        nom_tec = t['Tecnico']
                        tot_tec = t['Total_Tec']
                        pct_tec = (tot_tec / total_prog_unicas) * 100
                        
                        with st.expander(f"👨‍🔧 {nom_tec} — {tot_tec} OTs asignadas ({pct_tec:.1f}% de presencia)"):
                            
                            df_tec = df_prog_tec[df_prog_tec['Tecnico'] == nom_tec]
                            
                            # 💡 MAGIA AQUÍ: Agrupamos por Cliente y Sucursal al mismo tiempo
                            sucursales_tec = df_tec.groupby(['Cliente', 'Sucursal']).size().reset_index(name='Total_Suc').sort_values('Total_Suc', ascending=False)
                            
                            for _, c in sucursales_tec.iterrows():
                                nom_cli = c['Cliente']
                                nom_suc = c['Sucursal']
                                tot_suc = c['Total_Suc']
                                pct_suc = (tot_suc / tot_tec) * 100
                                
                                st.write(f"🏢 **{nom_cli}** | 📍 Sede: **{nom_suc}** ({pct_suc:.1f}% / {tot_suc} OTs)")
                                
                                df_suc_ots = df_tec[(df_tec['Cliente'] == nom_cli) & (df_tec['Sucursal'] == nom_suc)]
                                lista_formateada = []
                                
                                for _, fila_ot in df_suc_ots.iterrows():
                                    ot_val = str(fila_ot['OT_Num']).strip()
                                    eq_val = str(fila_ot['Equipo']).strip() if pd.notna(fila_ot['Equipo']) else "S/E"
                                    
                                    # Formato más limpio ahora que la sucursal ya está en el título
                                    lista_formateada.append(f"{ot_val} [Eq:{eq_val}]")
                                
                                texto_ots = ", ".join(lista_formateada)
                                st.caption(f"🏷️ **OTs:** `{texto_ots}`")
                                
                                st.progress(float(pct_suc) / 100.0)
                                st.write("")
                                
            st.markdown("---")
            with st.expander("👀 Ver documento temporal (Matriz de datos completa)"):
                st.dataframe(df_ots.sort_values(by=['Estado', 'Zona', 'Tecnico', 'Cliente', 'Sucursal']), hide_index=True, use_container_width=True)
                
        else:
            st.success("✅ No se encontraron OTs en estado 'Programada' o 'Aprobada'.")
            
    except Exception as e:
        st.error(f"⚠️ Error conectando a la BD. Detalle técnico: {e}")

# =====================================================================
# 📍 MÓDULO 5: BITÁCORA OPERATIVA POR VEHÍCULO (REPOSITORIO SAMM)
# =====================================================================
elif menu_seleccionado == "📍 5. Bitácora GPS":
   
    # =========================================================
    # 📝 DICCIONARIO MAESTRO DE FLOTA (Desde config.json)
    # =========================================================
    DICCIONARIO_FLOTA = CONFIG.get("diccionario_flota", {})

    # --- PERSISTENCIA DE SITIOS ---
    ARCHIVO_SITIOS = "sitios_interes.json"

    def cargar_json(ruta, default_val):
        if os.path.exists(ruta):
            try:
                with open(ruta, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return default_val
        return default_val

    def guardar_json(ruta, datos):
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=4)

    diccionario_sitios = cargar_json(ARCHIVO_SITIOS, {})

    for k, v in list(diccionario_sitios.items()):
        if isinstance(v, str):
            diccionario_sitios[k] = {"tipo": "texto", "patron": k}

    def calcular_distancia(lat1, lon1, lat2, lon2):
        try:
            lat1, lon1, lat2, lon2 = float(lat1), float(lon1), float(lat2), float(lon2)
            if any(pd.isna([lat1, lon1, lat2, lon2])): return 999999.0
            R = 6371000.0
            dlat = math.radians(lat2 - lat1)
            dlon = math.radians(lon2 - lon1)
            a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
            return R * (2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))
        except Exception:
            return 999999.0

    st.title("📍 Bitácora Logística de Travesías")
    st.markdown("Reconstrucción operativa automática basada en la base de datos central de telemetría.")

    # =========================================================
    # ⚙️ PANEL DE CONFIGURACIÓN (SOLO SITIOS Y GEOCERCAS)
    # =========================================================
    with st.expander("⚙️ Administrar Nombres Personalizados de Clientes / Sitios (Geocercas y Puntos)"):
        st.markdown("Nombra las coordenadas o direcciones del GPS.")
        tab_c, tab_t = st.tabs(["📍 Por Coordenada (Recomendado)", "📝 Por Texto"])
        
        with tab_c:
            c1, c2, c3, c4 = st.columns([2, 2.5, 1, 1.2])
            with c1: alias_c = st.text_input("Nombre Cliente/Sitio:", placeholder="Ej: DINO, CCL...", key="ac")
            with c2: coord_in = st.text_input("Pegar Coordenadas (De la tabla inferior):", placeholder="Ej: 10.9238, -74.8652", key="coord_in")
            with c3: radio_in = st.number_input("Radio (m):", value=200, step=50, key="r_in")
            with c4:
                st.write("")
                if st.button("➕ Guardar GPS", use_container_width=True):
                    if alias_c and coord_in:
                        try:
                            # Separa automáticamente lo que el usuario pegó
                            partes = coord_in.replace(" ", "").split(",")
                            if len(partes) == 2:
                                diccionario_sitios[alias_c.strip()] = {
                                    "tipo": "coordenada", 
                                    "lat": float(partes[0]), 
                                    "lon": float(partes[1]), 
                                    "radio": float(radio_in)
                                }
                                guardar_json(ARCHIVO_SITIOS, diccionario_sitios)
                                st.rerun()
                            else:
                                st.error("⚠️ Falta la coma entre latitud y longitud.")
                        except Exception:
                            st.error("⚠️ Formato inválido. Usa números separados por coma.")
        
        with tab_t:
            ct1, ct2, ct3 = st.columns([2, 2, 1])
            with ct1: pat_t = st.text_input("Texto del GPS:", key="pt")
            with ct2: ali_t = st.text_input("Nombre a mostrar:", key="at")
            with ct3:
                st.write("")
                if st.button("➕ Guardar Texto", use_container_width=True):
                    if pat_t and ali_t:
                        diccionario_sitios[ali_t.strip()] = {"tipo": "texto", "patron": pat_t.strip()}
                        guardar_json(ARCHIVO_SITIOS, diccionario_sitios)
                        st.rerun()
                    
        if diccionario_sitios:
            st.markdown("---")
            cols_sitios = st.columns(3)
            for idx, (n_sitio, cfg) in enumerate(diccionario_sitios.items()):
                with cols_sitios[idx % 3]:
                    tag = f"📍" if cfg.get('tipo') == 'coordenada' else f"📝"
                    if st.button(f"🗑️ {tag} {n_sitio}", key=f"del_sitio_{n_sitio}"):
                        del diccionario_sitios[n_sitio]
                        guardar_json(ARCHIVO_SITIOS, diccionario_sitios)
                        st.rerun()

   # =========================================================
    # 📁 GESTOR INTELIGENTE DE LECTURA Y AUTO-REFRESCO
    # =========================================================
    CARPETA_SAMM = "datos_samm"
    if not os.path.exists(CARPETA_SAMM): os.makedirs(CARPETA_SAMM)

    archivo_esperado = os.path.join(CARPETA_SAMM, "historial_gps.xlsx")
    fecha_hoy_str = datetime.now().strftime("%Y-%m-%d")

    col_f, col_b, col_i = st.columns([2, 2, 4])
    
    with col_f:
        fecha_seleccionada = st.date_input(
            "📅 Consultar historial por fecha:",
            value=datetime.now(),
            max_value=datetime.now(),
            key="fecha_gps_picker"
        )
    fecha_str = fecha_seleccionada.strftime("%Y-%m-%d")

    with col_b:
        st.write("") 
        # Solo usamos este botón si el usuario quiere forzar la API manualmente
        if st.button("🔄 Forzar Sincronización API", use_container_width=True):
            with st.spinner("Sincronizando con Guardian GPS..."):
                actualizar_historial_api(fecha_str)
                st.cache_data.clear()
                st.rerun()

    # LÓGICA DE AISLAMIENTO:
    # Si elegimos HOY, dejamos que el robot trabaje y auto-recargamos la web cada 5 mins.
    # Si elegimos el PASADO, forzamos la descarga por API (una sola vez) porque el robot no lo hará.
    if fecha_str != fecha_hoy_str:
        if st.session_state.get('ultima_fecha_descargada') != fecha_str or not os.path.exists(archivo_esperado):
            with st.spinner(f"Descargando histórico del {fecha_str}..."):
                actualizar_historial_api(fecha_str)
                st.session_state['ultima_fecha_descargada'] = fecha_str
                st.cache_data.clear()
    else:
        # Inyección de JavaScript para auto-refresco (300,000 ms = 5 minutos)
        # Permite que la web lea lo que el robot descarga sin clics manuales.
        import streamlit.components.v1 as components
        components.html("<script>setTimeout(function(){window.parent.location.reload();}, 300000);</script>", height=0)

    if not os.path.exists(archivo_esperado):
        st.warning(f"⚠️ El archivo maestro no existe. Verifica que el robot en segundo plano esté ejecutándose.")
        st.stop()

    timestamp_archivo = os.path.getmtime(archivo_esperado)
    fecha_modificacion = datetime.fromtimestamp(timestamp_archivo).strftime('%d/%m/%Y %I:%M %p')
    
    with col_i:
        st.write("")
        st.info(f"📂 **Base:** `{os.path.basename(archivo_esperado)}` | ⏱️ **Última escritura del robot:** `{fecha_modificacion}`")
        
    # =========================================================
    # 🚀 MOTOR DE CACHÉ DE EXCEL
    # =========================================================
    @st.cache_data(show_spinner="Leyendo y optimizando archivo GPS...")
    def cargar_datos_gps_optimizado(ruta, timestamp):
        try:
            df = pd.read_excel(ruta)
        except Exception:
            try:
                tablas = pd.read_html(ruta, header=0, encoding='utf-8')
                df = max(tablas, key=len)
            except Exception:
                return None, None, f"⚠️ El archivo `{ruta}` no es válido.", []
        
        df.columns = df.columns.astype(str).str.strip()
        if not any('veh' in str(c).lower() for c in df.columns):
            for idx, row in df.head(10).iterrows():
                if any('veh' in str(val).lower() for val in row.values):
                    df.columns = df.iloc[idx].astype(str).str.strip()
                    df = df.iloc[idx + 1:].reset_index(drop=True)
                    break
        
        columnas_reales = list(df.columns)
        col_v = next((c for c in columnas_reales if 'veh' in str(c).lower()), None)
        col_f = next((c for c in columnas_reales if 'fecha' in str(c).lower()), None)
        col_e = next((c for c in columnas_reales if 'evento' in str(c).lower()), None)
        col_la = next((c for c in columnas_reales if 'latitud' in str(c).lower()), None)
        col_lo = next((c for c in columnas_reales if 'longitud' in str(c).lower()), None)
        col_u = next((c for c in columnas_reales if 'ubicaci' in str(c).lower()), None)
        if not col_u: col_u = next((c for c in columnas_reales if 'direcci' in str(c).lower()), None)
        
        if not all([col_v, col_f, col_e, col_u, col_la, col_lo]):
            return None, None, "⚠️ El archivo no contiene todas las columnas requeridas (Falta Evento, Lat, Lon, etc.).", []

        df[col_f] = pd.to_datetime(df[col_f], errors='coerce')
        df = df.dropna(subset=[col_f])
        
        hora_inicio = pd.to_datetime("06:00", format='%H:%M').time()
        hora_fin = pd.to_datetime("19:00", format='%H:%M').time()
        df = df[(df[col_f].dt.time >= hora_inicio) & (df[col_f].dt.time <= hora_fin)]
        
        if df.empty:
            return None, None, "⚠️ No hay registros de GPS en horario operativo.", []
        
        df[col_la] = pd.to_numeric(df[col_la].astype(str).str.replace(',', '.'), errors='coerce')
        df[col_lo] = pd.to_numeric(df[col_lo].astype(str).str.replace(',', '.'), errors='coerce')
        df = df.sort_values(by=[col_v, col_f]).reset_index(drop=True)
        
        df_ign_global = df[df[col_e].astype(str).str.contains('Ignición', case=False, na=False)].copy()
        
        return df, df_ign_global, "", [col_v, col_f, col_e, col_u, col_la, col_lo]

    df_gps, df_ign_maestro, error_msg, cols_maestras = cargar_datos_gps_optimizado(archivo_esperado, timestamp_archivo)
    
    if error_msg:
        st.error(error_msg)
        st.stop()
        
    col_vehiculo, col_fecha, col_evento, col_ubicacion, col_lat, col_lon = cols_maestras

    # =========================================================
    # 🌍 MOTOR DE REVERSE GEOCODING (OPENSTREETMAP CACHEADO)
    # =========================================================
    ARCHIVO_CACHE_MAPAS = "datos_samm/cache_mapas.json"
    def cargar_cache_mapas():
        if os.path.exists(ARCHIVO_CACHE_MAPAS):
            try:
                with open(ARCHIVO_CACHE_MAPAS, "r", encoding="utf-8") as f: return json.load(f)
            except: return {}
        return {}

    cache_mapas = cargar_cache_mapas()
    mapas_modificados = False

    def obtener_calle_osm(lat, lon):
        global mapas_modificados
        # Redondear a 3 decimales (~110 metros) para agrupar puntos cercanos y optimizar llamadas
        lat_r, lon_r = round(float(lat), 3), round(float(lon), 3)
        llave = f"{lat_r},{lon_r}"
        
        if llave in cache_mapas: 
            return cache_mapas[llave]
            
        try:
            url = f"https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lon}&zoom=16"
            headers = {'User-Agent': 'DinoERP-Logistica/1.0'}
            resp = requests.get(url, headers=headers, timeout=2)
            if resp.status_code == 200:
                direccion_completa = resp.json().get('display_name', '')
                if direccion_completa:
                    # Nos quedamos con los 2 o 3 primeros segmentos para hacerlo legible (Ej: Calle 30, Soledad)
                    partes = [p.strip() for p in direccion_completa.split(',')]
                    texto_resumido = ", ".join(partes[:3]) if len(partes) >= 3 else direccion_completa
                    
                    cache_mapas[llave] = texto_resumido
                    mapas_modificados = True
                    time.sleep(0.2) # Pausa de respeto al servidor gratuito
                    return texto_resumido
        except: 
            pass
        return "Ubicación sin texto"

    def evaluar_direccion(val_txt, lat, lon):
        # 1. Prioridad Absoluta: Geocercas Locales (Clientes)
        if not pd.isna(lat) and not pd.isna(lon):
            for nombre_sitio, cfg in diccionario_sitios.items():
                if cfg.get("tipo") == "coordenada":
                    if calcular_distancia(lat, lon, cfg.get("lat"), cfg.get("lon")) <= cfg.get("radio", 200):
                        return f"🏢 {nombre_sitio}"

        txt = str(val_txt).strip()
        
        # 2. Traducción nativa de Guardian (Si el usuario sube un Excel descargado a mano)
        if txt not in ['-', 'nan', '', 'None']:
            for nombre_sitio, cfg in diccionario_sitios.items():
                if cfg.get("tipo") == "texto" and cfg.get("patron") and cfg.get("patron").lower() in txt.lower():
                    return f"🏢 {nombre_sitio}"
            txt = re.sub(r'^CO,[^,]+,', '', txt)
            match = re.search(r',en \((.*?)\)', txt)
            lugar_extra = f" ({match.group(1)})" if match else ""
            txt = re.sub(r' - Rel:[^,]+', '', txt)
            return re.sub(r',?en \(.*?\)', '', txt).replace(',', ', ') + lugar_extra

        # 3. Fallback Dinámico: Traducir Coordenada Cruda de la API a Texto
        if not pd.isna(lat) and not pd.isna(lon):
            return obtener_calle_osm(lat, lon)

        return "Ubicación sin texto"
    # =========================================================
    # 🚧 MOTOR DE PROCESAMIENTO Y FILTRO DE TIPOS
    # =========================================================
    vehiculos_unicos = sorted(df_gps[col_vehiculo].dropna().unique())
    
    tipos_disponibles = set()
    mapa_vehiculo_tipo = {}
    
    for v in vehiculos_unicos:
        placa = str(v).strip().upper()
        if placa in DICCIONARIO_FLOTA:
            tipo = DICCIONARIO_FLOTA[placa]["tipo"]
        else:
            tipo = "Otros"
            
        mapa_vehiculo_tipo[placa] = tipo
        tipos_disponibles.add(tipo)
        
    tipos_disponibles = sorted(list(tipos_disponibles))
    
    st.markdown("---")
    tipos_seleccionados = st.multiselect(
        "🎯 Filtrar por Tipo de Vehículo:",
        options=tipos_disponibles,
        default=tipos_disponibles
    )
    
    vehiculos_a_mostrar = [v for v in vehiculos_unicos if mapa_vehiculo_tipo[str(v).strip().upper()] in tipos_seleccionados]

    st.subheader(f"Flota Mostrada: {len(vehiculos_a_mostrar)} equipos")
    # Inicializar lista para el reporte global a exportar
    reporte_global = []
    
    for vehiculo in vehiculos_a_mostrar:
        vehiculo_str = str(vehiculo).strip().upper()
        
        info_vehiculo = DICCIONARIO_FLOTA.get(vehiculo_str, {"tipo": "Otros", "emoji": "🚗"})
        icono = info_vehiculo["emoji"]
        tipo_nombre = info_vehiculo["tipo"]

        df_veh = df_gps[df_gps[col_vehiculo] == vehiculo]
        df_ign = df_ign_maestro[df_ign_maestro[col_vehiculo] == vehiculo].copy().reset_index(drop=True)
        
        trayectos_vehiculo = []
        total_min_rodamiento = 0
        
        i = 0
        while i < len(df_ign):
            row = df_ign.iloc[i]
            evt_txt = str(row[col_evento])
            
            if 'prendida' in evt_txt.lower():
                origen_fecha = row[col_fecha]
                origen_dir = evaluar_direccion(row[col_ubicacion], row[col_lat], row[col_lon])
                
                if i + 1 < len(df_ign):
                    next_row = df_ign.iloc[i + 1]
                    next_evt = str(next_row[col_evento])
                    
                    if 'apagada' in next_evt.lower():
                        destino_fecha = next_row[col_fecha]
                        dest_lat, dest_lon = next_row[col_lat], next_row[col_lon]
                        destino_dir = evaluar_direccion(next_row[col_ubicacion], dest_lat, dest_lon)
                        
                        rodamiento_min = (destino_fecha - origen_fecha).total_seconds() / 60.0
                        total_min_rodamiento += rodamiento_min
                        
                       # 💡 CÁLCULO DE PERMANENCIA EN SITIO
                        if i + 2 < len(df_ign):
                            encendido_siguiente = df_ign.iloc[i + 2][col_fecha]
                            perm_min = (encendido_siguiente - destino_fecha).total_seconds() / 60.0
                            permanencia_str = f"{int(perm_min // 60)}h {int(perm_min % 60)}m" if perm_min >= 60 else f"{int(perm_min)} min"
                        else:
                            # 🏢 EVALUACIÓN DE RETORNO A BASE (DINO) SIN IMPORTAR LA HORA
                            es_dino = 'dino' in destino_dir.lower()
                            
                            if es_dino:
                                permanencia_str = "🏢 Guardado en dino"
                            else:
                                # Caso normal: terminó su último viaje en otro lado (ej. donde un cliente)
                                hora_actual = datetime.now()
                                fecha_ref = df_veh[col_fecha].max() if hora_actual.date() != destino_fecha.date() else hora_actual
                                perm_act_min = max(0, (fecha_ref - destino_fecha).total_seconds() / 60.0)
                                fmt_act = f"{int(perm_act_min // 60)}h {int(perm_act_min % 60)}m" if perm_act_min >= 60 else f"{int(perm_act_min)} min"
                                permanencia_str = f"🟢 En sitio desde {destino_fecha.strftime('%I:%M %p')} (Lleva {fmt_act})"
                        
                        fmt_rodamiento = f"{int(rodamiento_min // 60)}h {int(rodamiento_min % 60)}m" if rodamiento_min >= 60 else f"{int(rodamiento_min)} min"
                        
                        trayectos_vehiculo.append({
                            'Origen (Prendido)': f"{origen_dir} ({origen_fecha.strftime('%I:%M %p')})",
                            'Rodamiento': fmt_rodamiento,
                            'Destino (Apagado)': f"{destino_dir} ({destino_fecha.strftime('%I:%M %p')})",
                            'Permanencia en Sitio': permanencia_str,
                            'Coordenadas (Copiar)': f"{dest_lat}, {dest_lon}",
                            'Mapa Destino': f"https://www.google.com/maps?q={dest_lat},{dest_lon}"
                        })
                        
                        # 💡 NUEVO: Agregar al reporte global para el Director
                        reporte_global.append({
                            'Tipo': tipo_nombre,
                            'Vehículo': vehiculo_str,
                            'Origen': f"{origen_dir} ({origen_fecha.strftime('%I:%M %p')})",
                            'Destino': f"{destino_dir} ({destino_fecha.strftime('%I:%M %p')})",
                            'Rodamiento': fmt_rodamiento,
                            'Permanencia': permanencia_str,
                            'Link Mapa': f"https://www.google.com/maps?q={dest_lat},{dest_lon}"
                        })
                        
                        i += 1
            i += 1

        cant_trayectos = len(trayectos_vehiculo)
        ultimo_dest = trayectos_vehiculo[-1]['Destino (Apagado)'] if cant_trayectos > 0 else "Sin registros completados"
        
        titulo_expander = f"{icono} **{tipo_nombre}: {vehiculo_str}** | {cant_trayectos} Trayectos | Último: {ultimo_dest}"
        
        with st.expander(titulo_expander, expanded=False):
            if cant_trayectos > 0:
                df_veh_tabla = pd.DataFrame(trayectos_vehiculo)
                col_m1, col_m2 = st.columns(2)
                col_m1.metric("Total Trayectos", f"{cant_trayectos} viajes")
                fmt_total_rod = f"{int(total_min_rodamiento // 60)}h {int(total_min_rodamiento % 60)}m" if total_min_rodamiento >= 60 else f"{int(total_min_rodamiento)} min"
                col_m2.metric("Tiempo Total en Carretera", fmt_total_rod)
                
                st.dataframe(df_veh_tabla, use_container_width=True, hide_index=True, column_config={"Mapa Destino": st.column_config.LinkColumn("📍 Mapa", display_text="Ver Destino")})
            else:
                st.info(f"No hay trayectos cerrados para {tipo_nombre.lower()} hoy.")

    # Guardar en disco las nuevas calles descubiertas por OSM
    if mapas_modificados:
        with open(ARCHIVO_CACHE_MAPAS, "w", encoding="utf-8") as f:
            json.dump(cache_mapas, f, ensure_ascii=False, indent=4)

    # =========================================================
    # 📥 BOTONES DE EXPORTACIÓN DE REPORTE EJECUTIVO (EXCEL Y PDF)
    # =========================================================
    if reporte_global:
        st.markdown("---")
        st.subheader("📄 Reporte Consolidado de la Jornada")
        st.markdown("Descarga el informe completo con las ubicaciones reconstruidas, tiempos y links interactivos.")
        
        col_btn1, col_btn2 = st.columns(2)
        
        # Generar DataFrame Base
        df_export = pd.DataFrame(reporte_global)
        fecha_str_hoy = datetime.now().strftime('%d_%m_%Y')
        
        # ---------------------------------------------------------
        # 1. BOTÓN PARA EXCEL (CORREGIDO)
        # ---------------------------------------------------------
        import io
        buffer_excel = io.BytesIO()
        try:
            with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Rutas_Completadas')
                worksheet = writer.sheets['Rutas_Completadas']
                for idx, col in enumerate(df_export.columns):
                    max_len = max(df_export[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(idx, idx, max_len)
                    
            with col_btn1:
                # ERROR CORREGIDO AQUÍ: Era download_but75ton
                st.download_button(
                    label="📊 Descargar Ejecutivo (Excel)",
                    data=buffer_excel.getvalue(),
                    file_name=f"Reporte_Rutas_{fecha_str_hoy}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        except Exception as e:
            col_btn1.error("Falta xlsxwriter (pip install xlsxwriter)")

        # ---------------------------------------------------------
        # 2. BOTÓN PARA PDF (CON CÁPSULAS INTELIGENTES Y LOGO GRANDE)
        # ---------------------------------------------------------
        with col_btn2:
            try:
                from fpdf import FPDF
                
                # 1. Ajuste del Logo en el Lienzo
                class ReporteGPS_Dino(FPDF):
                    def header(self):
                        azul_dino = (0, 32, 96)
                        ruta_logo = "logo.png"
                        if os.path.exists(ruta_logo):
                            self.image(ruta_logo, x=15, y=12, w=75) 
                        
                        self.set_font("helvetica", "B", 14)
                        self.set_text_color(*azul_dino)
                        self.set_y(15)
                        self.cell(0, 6, "REPORTE OPERATIVO DE RUTAS", ln=True, align='C')
                        
                        self.set_font("helvetica", "I", 10)
                        self.set_text_color(100, 100, 100)
                        self.cell(0, 6, f"Fecha de reporte: {fecha_str_hoy}", ln=True, align='C')
                        self.ln(12)
                        
                    def footer(self):
                        self.set_y(-15)
                        self.set_font("helvetica", "I", 8)
                        self.set_text_color(128, 128, 128)
                        self.cell(0, 10, f"Página {self.page_no()}", align="C")

                pdf = ReporteGPS_Dino(orientation='L', unit='mm', format='A4')
                pdf.add_page()
                
                azul_dino = (0, 32, 96)
                naranja_dino = (255, 102, 0)
                gris_borde = (200, 200, 200)
                
                vehiculos_en_reporte = df_export['Vehículo'].unique()
                
                for veh in vehiculos_en_reporte:
                    df_veh = df_export[df_export['Vehículo'] == veh]
                    tipo_veh = df_veh.iloc[0]['Tipo']
                    total_viajes = len(df_veh)
                    
                    # ⏱️ Calcular rodamiento total del vehículo reconstruyendo los strings
                    total_mins = 0
                    for r_str in df_veh['Rodamiento']:
                        try:
                            if 'h' in str(r_str):
                                partes = str(r_str).split('h')
                                h = int(partes[0].strip())
                                m_str = partes[1].replace('m', '').replace('in', '').strip()
                                m = int(m_str) if m_str else 0
                                total_mins += h * 60 + m
                            else:
                                m = int(str(r_str).replace('min', '').strip())
                                total_mins += m
                        except:
                            pass
                    
                    fmt_total_rod = f"{int(total_mins // 60)}h {int(total_mins % 60)}m" if total_mins >= 60 else f"{int(total_mins)} min"
                    
                    # 🔥 MOTOR DE PAGINACIÓN INTELIGENTE (CHUNKING)
                    df_restante = df_veh.copy()
                    es_primer_bloque = True
                    
                    while not df_restante.empty:
                        # Evaluar si estamos muy abajo en la hoja
                        if pdf.get_y() > 165: 
                            pdf.add_page()
                            
                        pdf.ln(4)
                        block_y = pdf.get_y()
                        
                        espacio_disponible = 185 - block_y
                        
                        if es_primer_bloque:
                            alto_encabezado = 27
                        else:
                            alto_encabezado = 7
                            
                        espacio_para_filas = espacio_disponible - alto_encabezado
                        filas_que_caben = int(espacio_para_filas // 7)
                        
                        if filas_que_caben < 1: 
                            pdf.add_page()
                            continue
                            
                        chunk = df_restante.iloc[:filas_que_caben]
                        df_restante = df_restante.iloc[filas_que_caben:] 
                        
                        # 1. Dibujar la cápsula contenedora perfecta
                        altura_tarjeta = alto_encabezado + (len(chunk) * 7)
                        pdf.set_draw_color(*gris_borde)
                        pdf.set_line_width(0.3)
                        
                        try:
                            pdf.rect(x=10, y=block_y, w=277, h=altura_tarjeta, style='D', round_corners=True, corner_radius=3)
                        except:
                            pdf.rect(x=10, y=block_y, w=277, h=altura_tarjeta, style='D')

                        # 2. Encabezado de la Tarjeta
                        if es_primer_bloque:
                            pdf.set_draw_color(*naranja_dino)
                            pdf.set_line_width(1.5)
                            pdf.line(13, block_y + 4, 13, block_y + 14)

                            pdf.set_xy(16, block_y + 3.5)
                            pdf.set_font("helvetica", "B", 12)
                            pdf.set_text_color(*azul_dino)
                            
                            pdf.cell(0, 6, f"EQUIPO: {veh}   |   TIPO: {tipo_veh.upper()}   |   TIEMPO EN RUTA: {fmt_total_rod}", ln=True)
                            
                            pdf.set_draw_color(230, 230, 230)
                            pdf.set_line_width(0.2)
                            pdf.line(16, block_y + 10.5, 280, block_y + 10.5)

                            pdf.set_xy(16, block_y + 11.5)
                            pdf.set_font("helvetica", "B", 9)
                            pdf.set_text_color(*naranja_dino)
                            pdf.cell(0, 5, f"Operación del día: {total_viajes} rutas registradas", ln=True)
                            
                            table_y = block_y + 20
                            pdf.set_draw_color(*gris_borde)
                            pdf.set_line_width(0.1)
                            pdf.line(10, table_y, 287, table_y)
                        else:
                            table_y = block_y
                            
                        # 3. Cabeceras de Tabla
                        pdf.set_xy(10, table_y)
                        pdf.set_draw_color(*gris_borde)
                        pdf.set_line_width(0.1)
                        pdf.set_fill_color(245, 245, 245)
                        pdf.set_font("helvetica", "B", 8)
                        pdf.set_text_color(*azul_dino)
                        
                        col_w = [85, 85, 42, 42, 23] 
                        
                        pdf.cell(col_w[0], 7, "ORIGEN (PRENDIDO)", border='R', fill=True, align='C')
                        pdf.cell(col_w[1], 7, "DESTINO (APAGADO)", border='R', fill=True, align='C')
                        pdf.cell(col_w[2], 7, "RODAMIENTO", border='R', fill=True, align='C')
                        pdf.cell(col_w[3], 7, "PERMANENCIA", border='R', fill=True, align='C')
                        pdf.cell(col_w[4], 7, "MAPA", border=0, fill=True, align='C')
                        pdf.ln()
                        pdf.line(10, pdf.get_y(), 287, pdf.get_y())

                        # 4. Filas de datos del bloque actual
                        pdf.set_font("helvetica", "", 7)
                        pdf.set_text_color(60, 60, 60)
                        
                        for idx_row, (_, fila) in enumerate(chunk.iterrows()):
                            origen = str(fila['Origen'])
                            destino = str(fila['Destino'])
                            rodamiento = str(fila['Rodamiento'])
                            permanencia = str(fila['Permanencia'])
                            
                            permanencia = permanencia.replace("🟢", ">>")
                            if len(origen) > 58: origen = origen[:55] + "..."
                            if len(destino) > 58: destino = destino[:55] + "..."
                            
                            origen = origen.encode('latin-1', 'ignore').decode('latin-1')
                            destino = destino.encode('latin-1', 'ignore').decode('latin-1')
                            rodamiento = rodamiento.encode('latin-1', 'ignore').decode('latin-1')
                            permanencia = permanencia.encode('latin-1', 'ignore').decode('latin-1')
                            
                            pdf.cell(col_w[0], 7, f"  {origen}", border='R', align='L')
                            pdf.cell(col_w[1], 7, f"  {destino}", border='R', align='L')
                            pdf.cell(col_w[2], 7, rodamiento, border='R', align='C')
                            pdf.cell(col_w[3], 7, permanencia, border='R', align='C')
                            
                            pdf.set_text_color(*naranja_dino)
                            pdf.set_font("helvetica", "U", 7)
                            pdf.cell(col_w[4], 7, "Ver Mapa", border=0, align='C', link=fila['Link Mapa'])
                            
                            pdf.set_text_color(60, 60, 60)
                            pdf.set_font("helvetica", "", 7)
                            pdf.ln()
                            
                            if idx_row < len(chunk) - 1:
                                pdf.line(10, pdf.get_y(), 287, pdf.get_y())

                        pdf.set_y(block_y + altura_tarjeta)
                        es_primer_bloque = False

                # 5. Generar archivo blindado
                resultado_pdf = pdf.output(dest='S')
                if isinstance(resultado_pdf, str):
                    pdf_bytes = resultado_pdf.encode('latin-1')
                else:
                    pdf_bytes = bytes(resultado_pdf)
                
                st.download_button(
                    label="📄 Descargar Ejecutivo (PDF)",
                    data=pdf_bytes,
                    file_name=f"Reporte_Rutas_{fecha_str_hoy}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ Error construyendo el PDF con FPDF: {e}")


# =====================================================================
# 🗂️ MÓDULO 6: TRAZABILIDAD DOCUMENTAL (OT -> RQ -> OC)
# =====================================================================
elif menu_seleccionado == "📝 6. Reportes y Descargas":
    st.title("🗂️ Trazabilidad Documental de Equipos")
    st.markdown("Consulta el árbol de documentos (Órdenes de Trabajo, Requisiciones y Órdenes de Compra) asociados a un equipo.")
    st.markdown("---")
    
    # Extraemos la lista de equipos activos desde la base maestra
    lista_equipos_historico = ["Seleccionar..."] + sorted(st.session_state['df_base_maestra']['equipo'].dropna().unique().tolist())
    
    col_busqueda, col_vacia = st.columns([1, 2])
    with col_busqueda:
        equipo_consulta = st.selectbox("🔍 Selecciona el número del equipo:", lista_equipos_historico)
        btn_consultar = st.button("Buscar Historial", type="primary", use_container_width=True)

    if btn_consultar and equipo_consulta != "Seleccionar...":
        with st.spinner(f"Rastreando árbol documental en SAMM para el equipo {equipo_consulta}..."):
            try:
                conn = st.connection("sw_dino", type="sql")
                
                # Consulta SQL ajustada a la estructura real de SAMM (Prefijo + Número)
                query_trazabilidad = f"""
                    SELECT 
                        ISNULL(ot_base.prefijo, 'OTT') + ' ' + CAST(ot_base.documento_numero AS VARCHAR) AS [Orden de Trabajo (OT)],
                        CONVERT(varchar, ot_base.fecha_fh, 103) AS [Fecha],
                        ISNULL(STRING_AGG(ISNULL(rq.prefijo, 'RQ') + ' ' + CAST(rq.documento_numero AS VARCHAR), ' | '), 'N/A') AS [Requisiciones (RQ)],
                        ISNULL(STRING_AGG(ISNULL(oc.prefijo, 'OC') + ' ' + CAST(oc.documento_numero AS VARCHAR), ' | '), 'N/A') AS [Ordenes de Compra (OC)]
                    FROM view_doc_documento_ot AS ot_view
                    JOIN doc_documento AS ot_base ON ot_view.id = ot_base.id
                    
                    -- Buscamos hijos (Requisiciones) por su prefijo
                    LEFT JOIN doc_documento AS rq 
                        ON rq.id_documento = ot_base.id 
                       AND rq.prefijo IN ('RQ', 'REQ', 'R.Q')
                        
                    -- Buscamos hijos (Ordenes de Compra) ligados a la RQ o directo a la OT
                    LEFT JOIN doc_documento AS oc 
                        ON (oc.id_documento = rq.id OR oc.id_documento = ot_base.id) 
                       AND oc.prefijo IN ('OC', 'O.C')
                        
                    -- Limpiamos espacios invisibles para traer las 14 OTs exactas
                    WHERE LTRIM(RTRIM(ot_view.equ_equipo_equipo)) = LTRIM(RTRIM('{equipo_consulta}'))
                    GROUP BY 
                        ot_base.prefijo,
                        ot_base.documento_numero,
                        ot_base.fecha_fh
                    ORDER BY ot_base.fecha_fh DESC
                """
                
                df_historial = conn.query(query_trazabilidad, ttl=0) 
                
                if not df_historial.empty:
                    st.success(f"✅ Se encontraron {len(df_historial)} Órdenes de Trabajo para el equipo {equipo_consulta}.")
                    
                    st.dataframe(
                        df_historial, 
                        use_container_width=True, 
                        hide_index=True
                    )
                    
                    import io
                    buffer_csv = io.BytesIO()
                    df_historial.to_csv(buffer_csv, index=False, sep=';', encoding='utf-8-sig')
                    st.download_button(
                        label="📥 Descargar Historial (CSV)",
                        data=buffer_csv.getvalue(),
                        file_name=f"Trazabilidad_{equipo_consulta}.csv",
                        mime="text/csv"
                    )
                else:
                    st.warning(f"No se encontraron registros de OTs para el equipo {equipo_consulta}.")
                    
            except Exception as e:
                st.error(f"⚠️ Error al consultar la base de datos: {e}")
