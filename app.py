import streamlit as st
import pandas as pd
import io
import os
import glob
import re
from modulos.parser_samm import limpiar_reporte_samm
from modulos.generador_excel import generar_pdf_cliente, generar_pdf_interno_dino

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Dinomontacargas - ERP Mantenimiento", 
    page_icon="🚜", 
    layout="wide"
)

# --- MENÚ LATERAL (ENRUTADOR) ---
with st.sidebar:
    st.header("🛠️ Módulos de Gestión")
    menu_seleccionado = st.radio(
        "Navegación:",
        [
            "📅 1. Gestión de Cronogramas", 
            "🛢️ 2. Predictivo de Horómetros", 
            "🚜 3. Directorio de Flota"
        ]
    )
    
    st.markdown("---")
    st.header("📥 Ingreso de Datos (Módulo 1)")
    st.info("Sube el reporte de SAMM (Detalle_Visitas).")
    archivo_samm = st.file_uploader("Cargar SAMM", type=["xls", "xlsx"], key="samm")

# =====================================================================
# --- CEREBRO GLOBAL: LECTURA DE LA BASE MAESTRA ---
# =====================================================================
if 'df_base_maestra' not in st.session_state:
    if not os.path.exists("datos_samm"):
        st.error("⚠️ La carpeta 'datos_samm' no existe. Por favor, créala y mete tu base de equipos ahí.")
        st.stop()
        
    archivos = glob.glob("datos_samm/*.xls*")
    archivos_maestros = [f for f in archivos if "Control_Mantenimiento" not in f]
    
    if not archivos_maestros:
        st.error("⚠️ La carpeta 'datos_samm' está vacía. Coloca tu archivo de equipos ahí.")
        st.stop()
        
    try:
        ruta_maestro = archivos_maestros[0]
        df_maestro = pd.read_excel(ruta_maestro) 
        st.session_state['df_base_maestra'] = df_maestro
    except Exception as e:
        st.error(f"⚠️ Error crítico al leer la Base Maestra. Detalle: {e}")
        st.stop()

# =====================================================================
# 🟩 MÓDULO 1: GESTIÓN DE CRONOGRAMAS 
# =====================================================================
if menu_seleccionado == "📅 1. Gestión de Cronogramas":
    st.title("📅 Gestión de Cronogramas y Auditoría")
    st.markdown("---")
    
    if archivo_samm is not None:
        try:
            df_maestro_actual = st.session_state['df_base_maestra']
            dict_tercero = dict(zip(df_maestro_actual['equipo'].astype(str).str.strip(), df_maestro_actual['Tercero']))
            dict_sucursal = dict(zip(df_maestro_actual['equipo'].astype(str).str.strip(), df_maestro_actual['Sucursal']))

            if 'df_master' not in st.session_state or st.session_state.get('nombre_archivo') != archivo_samm.name:
                df_crudo = limpiar_reporte_samm(archivo_samm)
                
                def normalizar_id_auditoria(val):
                    import re
                    val_str = str(val)
                    match = re.search(r'\[\s*(\d+)\s*\]', val_str)
                    if match: return match.group(1)
                    val_str = val_str.strip()
                    if val_str.endswith('.0'): return val_str[:-2]
                    return val_str

                def auditar_contra_maestro(row):
                    equipo = normalizar_id_auditoria(row['Equipo'])
                    cliente_samm = str(row['Cliente']).upper().strip()
                    
                    # Validación PROFESIONAL de nulos
                    tercero_real_raw = dict_tercero.get(equipo)
                    if pd.isna(tercero_real_raw) or str(tercero_real_raw).strip() == "" or str(tercero_real_raw).strip().upper() == "NAN": 
                        return "VERDE" 
                        
                    tercero_real = str(tercero_real_raw).upper().strip()
                    if cliente_samm not in tercero_real and tercero_real not in cliente_samm: 
                        return "ROJO" 
                    return "VERDE"
                    
                df_crudo['Alerta_Auditoria'] = df_crudo.apply(auditar_contra_maestro, axis=1)

                # ========================================================
                # NUEVO: MOTOR DE AGRUPACIÓN LOGÍSTICA (CLUSTERING 3 DÍAS)
                # ========================================================
                def optimizar_fechas_por_sucursal(df):
                    df_opt = df.copy()
                    df_opt['Fecha_DT'] = pd.to_datetime(df_opt['Fecha_Visita'].astype(str).str.split(" ").str[0], dayfirst=True, errors='coerce')
                    
                    cambios_realizados = 0
                    grupos = df_opt.groupby(['Cliente', 'Sucursal'])
                    
                    for (cliente, sucursal), grupo in grupos:
                        if pd.isna(sucursal) or str(sucursal).strip() == "" or str(sucursal).upper() == "NAN": continue
                        
                        if len(grupo) > 1:
                            grupo_ordenado = grupo.sort_values('Fecha_DT')
                            fecha_base = None
                            fecha_base_str = None
                            
                            for idx, row in grupo_ordenado.iterrows():
                                if pd.isna(row['Fecha_DT']): continue
                                
                                if fecha_base is None:
                                    fecha_base = row['Fecha_DT']
                                    fecha_base_str = str(df_opt.at[idx, 'Fecha_Visita'])
                                    continue
                                    
                                diferencia = (row['Fecha_DT'] - fecha_base).days
                                
                                if 0 < diferencia <= 3:
                                    df_opt.at[idx, 'Fecha_Visita'] = fecha_base_str
                                    cambios_realizados += 1
                                elif diferencia > 3:
                                    fecha_base = row['Fecha_DT']
                                    fecha_base_str = str(df_opt.at[idx, 'Fecha_Visita'])
                                    
                    df_opt = df_opt.drop(columns=['Fecha_DT'])
                    return df_opt, cambios_realizados

                df_crudo, total_optimizados = optimizar_fechas_por_sucursal(df_crudo)
                
                if total_optimizados > 0:
                    st.toast(f"Clustering Activo: Se unificaron automáticamente {total_optimizados} visitas en las mismas sucursales (Margen: 3 días).", icon="🚜")

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
            
            tab_buscador, tab_datos, tab_cronograma = st.tabs([
                "🔍 Buscador de Flota", "📋 Base de Datos General", "📅 Generar Cronogramas y Auditoría"
            ])
            
            with tab_buscador:
                st.subheader("Radiografía por Equipo")
                lista_equipos = sorted(df_limpio['Equipo'].astype(str).unique())
                equipo_buscado = st.selectbox("Selecciona el Equipo:", options=["Seleccionar..."] + lista_equipos)
                if equipo_buscado != "Seleccionar...":
                    df_equipo = df_limpio[df_limpio['Equipo'].astype(str) == equipo_buscado].sort_values(by="Fecha_Visita")
                    c1, c2, c3 = st.columns(3)
                    c1.info(f"**🏢 Cliente:**\n{df_equipo['Cliente'].iloc[0]}")
                    c2.info(f"**📍 Ubicación:**\n{df_equipo['Sucursal'].iloc[0]}")
                    c3.info(f"**🔧 Visitas:**\n{len(df_equipo)}")
                    st.dataframe(df_equipo[['Fecha_Visita', 'Mantenimiento', 'OT', 'Estado']], use_container_width=True, hide_index=True)
            
            with tab_datos:
                st.subheader("Auditoría Global de la Flota (Filtro para Comercial)")
                df_anomalias = df_limpio[df_limpio['Alerta_Auditoria'] == 'ROJO'].drop_duplicates(subset=['Equipo', 'Cliente', 'Sucursal']).copy()
                
                if not df_anomalias.empty:
                    df_reporte_base = df_anomalias[['Equipo', 'Cliente', 'Sucursal', 'Estado', 'Mantenimiento']].rename(columns={'Cliente': 'Contrato en SAMM', 'Sucursal': 'Ubicación Física Real'})
                    df_reporte_base.insert(0, '✅ Es Sucursal Válida', False)
                    df_editado = st.data_editor(df_reporte_base, hide_index=True, use_container_width=True, disabled=['Equipo', 'Contrato en SAMM', 'Ubicación Física Real', 'Estado', 'Mantenimiento'])
                    
                    df_errores_reales = df_editado[df_editado['✅ Es Sucursal Válida'] == False].drop(columns=['✅ Es Sucursal Válida'])
                    col_a, col_b = st.columns([2, 1])
                    col_a.warning(f"Exportando {len(df_errores_reales)} errores.")
                    with col_b:
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            df_errores_reales.to_excel(writer, sheet_name='Anomalias', index=False)
                        st.download_button("📥 Descargar Reporte Depurado", buffer.getvalue(), "Anomalias.xlsx", "application/vnd.ms-excel", type="primary")
                else:
                    st.success("✅ No hay anomalías de ubicación.")
                    
                st.write("---")
    
                st.subheader("🚨 Equipos Sin Cronograma de Mantenimiento")
                df_maestro = st.session_state['df_base_maestra']
                
                def normalizar_id(val):
                    val_str = str(val)
                    match = re.search(r'\[\s*(\d+)\s*\]', val_str)
                    if match: return match.group(1)
                    val_str = val_str.strip()
                    if val_str.endswith('.0'): return val_str[:-2]
                    return val_str
                
                equipos_en_samm = df_limpio['Equipo'].apply(normalizar_id).unique()
                df_maestro['equipo_str'] = df_maestro['equipo'].apply(normalizar_id)
                df_faltantes = df_maestro[~df_maestro['equipo_str'].isin(equipos_en_samm)]
                df_faltantes_mostrar = df_faltantes[['equipo', 'Tercero', 'Sucursal', 'Modelo', 'Horometro Actual']].dropna(subset=['Tercero'])
                
                if not df_faltantes_mostrar.empty:
                    st.warning(f"{len(df_faltantes_mostrar)} equipos sin visita programada.")
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
                    
                    # =========================================================
                    # PROGRAMADOR MANUAL PARA EQUIPOS HUÉRFANOS
                    # =========================================================
                    st.write("---")
                    with st.expander("🛠️ Asignar a un Cronograma (Modo Manual Temporal)"):
                        st.info("Inyecta equipos huérfanos directamente al cronograma de un cliente en memoria.")
                        
                        col_eq, col_cli = st.columns(2)
                        lista_huerfanos = sorted(df_faltantes_mostrar['equipo'].astype(str).unique())
                        
                        eq_manual = col_eq.selectbox("1. Equipo sin proyección:", lista_huerfanos)
                        
                        clientes_existentes = sorted(df_limpio['Cliente'].unique())
                        tipo_cli = col_cli.radio("2. Asignar a:", ["Cliente Existente", "Crear Nuevo Cliente"])
                        if tipo_cli == "Cliente Existente":
                            cli_destino = col_cli.selectbox("Selecciona el cliente:", clientes_existentes)
                        else:
                            cli_destino = col_cli.text_input("Escribe el nuevo cliente:")
                            
                        st.markdown("**3. Fechas de Visita (Formato: DD/MM/YYYY)**")
                        num_visitas = st.number_input("¿Cuántas visitas deseas programar?", min_value=1, max_value=6, value=1)
                        
                        cols_fechas = st.columns(num_visitas)
                        fechas_ingresadas = []
                        
                        for i in range(num_visitas):
                            label = f"Visita {i+1} {'*' if i == 0 else ''}"
                            f = cols_fechas[i].text_input(label, key=f"fecha_manual_{i}", placeholder="Ej: 15/08/2026")
                            fechas_ingresadas.append(f)
                            
                        if st.button("💾 Guardar y Asignar Equipo", type="primary"):
                            if not cli_destino or cli_destino.strip() == "":
                                st.error("⚠️ El nombre del cliente es obligatorio.")
                            elif not fechas_ingresadas[0] or fechas_ingresadas[0].strip() == "":
                                st.error("⚠️ Debes ingresar obligatoriamente la primera fecha.")
                            else:
                                info_eq = df_faltantes_mostrar[df_faltantes_mostrar['equipo'].astype(str) == eq_manual].iloc[0]
                                sucursal_eq = info_eq['Sucursal'] if pd.notna(info_eq['Sucursal']) else "SIN SUCURSAL"
                                
                                nuevas_filas = []
                                for fecha in fechas_ingresadas:
                                    if fecha and fecha.strip() != "":
                                        nueva_fila = {
                                            'Equipo': eq_manual,
                                            'Cliente': cli_destino.upper().strip(),
                                            'Sucursal': sucursal_eq,
                                            'Fecha_Visita': fecha.strip(), 
                                            'Estado': 'PROGRAMADO MANUAL',
                                            'Mantenimiento': 'PREVENTIVO',
                                            'Color_Semantico': 'Amarillo',
                                            'Alerta_Auditoria': 'VERDE'
                                        }
                                        nuevas_filas.append(nueva_fila)
                                
                                df_nuevas = pd.DataFrame(nuevas_filas)
                                st.session_state['df_master'] = pd.concat([st.session_state['df_master'], df_nuevas], ignore_index=True)
                                
                                st.success(f"✅ ¡Equipo {eq_manual} inyectado exitosamente al contrato {cli_destino.upper()}!")
                                st.rerun()
                else:
                    st.success("✅ Todos tienen cronograma.")

            with tab_cronograma:
                st.subheader("Auditoría y Automatización de PDF (Clientes)")
                clientes_unicos = sorted(df_limpio['Cliente'].unique())
                opciones_selector = [f"👀 {c}" if 'ROJO' in df_limpio[df_limpio['Cliente'] == c]['Alerta_Auditoria'].values else c for c in clientes_unicos]
                
                cliente_seleccionado = st.selectbox("Cliente:", opciones_selector).replace("👀 ", "")
                df_filtrado = df_limpio[df_limpio['Cliente'] == cliente_seleccionado]
                alertas_rojas = df_filtrado[df_filtrado['Alerta_Auditoria'] == 'ROJO']
                
                if not alertas_rojas.empty:
                    st.warning("👀 REVISIÓN SUGERIDA: Ubicaciones erróneas detectadas en este cliente.")
                    st.dataframe(alertas_rojas[['Equipo', 'Cliente', 'Sucursal']].drop_duplicates(), hide_index=True)
                    
                # --- NUEVO: CREADOR DE CONTRATOS FANTASMA Y REASIGNACIÓN ---
                with st.expander("🛠️ Reasignar Equipo o Crear Nuevo Contrato (Manual)"):
                    st.info("💡 Mueve un equipo a otro cliente o crea un contrato temporal para agrupar equipos sin contrato en SAMM.")
                    c_eq, c_dst = st.columns(2)
                    
                    todos_los_equipos = sorted(df_limpio['Equipo'].unique())
                    eq_mover = c_eq.selectbox("1. Equipo a mover:", todos_los_equipos)
                    
                    tipo_dest = c_dst.radio("2. Tipo de asignación:", ["A un Cliente Existente", "Crear NUEVO Contrato"])
                    
                    if tipo_dest == "A un Cliente Existente":
                        dst = c_dst.selectbox("3. Selecciona el destino:", clientes_unicos)
                    else:
                        dst = c_dst.text_input("3. Escribe el nombre del nuevo contrato:")
                        
                    if st.button("🔄 Aplicar Cambio y Mover Equipo", type="primary"):
                        if dst and dst.strip() != "":
                            nuevo_cliente = dst.strip().upper()
                            st.session_state['df_master'].loc[st.session_state['df_master']['Equipo'] == eq_mover, 'Cliente'] = nuevo_cliente
                            
                            def re_auditar(row):
                                # 1. Traductor Inyectado
                                import re
                                val_str = str(row['Equipo'])
                                match = re.search(r'\[\s*(\d+)\s*\]', val_str)
                                eq_s = match.group(1) if match else val_str.strip()
                                if eq_s.endswith('.0'): eq_s = eq_s[:-2]
                                
                                # 2. Validación PROFESIONAL de nulos
                                c_samm = str(row['Cliente']).upper().strip()
                                tercero_real_raw = dict_tercero.get(eq_s)
                                
                                if pd.isna(tercero_real_raw) or str(tercero_real_raw).strip() == "" or str(tercero_real_raw).strip().upper() == "NAN":
                                    return "VERDE"
                                    
                                t_real = str(tercero_real_raw).upper().strip()
                                if c_samm not in t_real and t_real not in c_samm: 
                                    return "ROJO"
                                return "VERDE"
                                
                            # Re-evaluamos toda la flota y refrescamos la pantalla
                            st.session_state['df_master']['Alerta_Auditoria'] = st.session_state['df_master'].apply(re_auditar, axis=1)
                            st.rerun()
                        else:
                            st.error("⚠️ Debes especificar un destino válido.")
                            
                st.write("---")
                
                # --- TEXTO PERSONALIZADO Y LISTA DE COMBUSTIÓN ---
                st.subheader("📝 Novedades y Observaciones del Mes")
                texto_novedad = st.text_area(
                    "Agrega una nota personalizada para este cliente (opcional):", 
                    placeholder="Ej: Durante este mes se observó un desgaste irregular en las llantas del equipo 408..."
                )
                
                # Leemos la lista de combustión para saber de qué tipo es cada equipo
                lista_combustion = []
                archivo_tracker = "datos_samm/Control_Mantenimiento.xlsx"
                if os.path.exists(archivo_tracker):
                    try:
                        df_tr = pd.read_excel(archivo_tracker)
                        col = 'EQUIPO' if 'EQUIPO' in df_tr.columns else 'Equipo'
                        lista_combustion = df_tr[col].astype(str).str.strip().tolist()
                    except: pass
                # ---------------------------------------------------------

                # Ahora le pasamos la novedad y la lista de equipos a la función del PDF
                pdf_bytes = generar_pdf_cliente(df_filtrado, cliente_seleccionado, texto_novedad, lista_combustion)
                
                st.download_button(f"⬇️ Descargar PDF - {cliente_seleccionado}", pdf_bytes, f"Cronograma_{cliente_seleccionado}.pdf", "application/pdf", type="primary")

                # ==========================================================
                # PLANIFICADOR LOGÍSTICO INTERNO (RUTA DINO)
                # ==========================================================
                st.markdown("---")
                st.subheader("🚜 Planificador Logístico Interno (Mantenimiento Dino)")
                st.info("Calcula las horas técnicas del mes, evaluando restricciones FÍSICAS (Sucursal).")
                
                # 1. Filtro Inverso (Blacklist): Zonas o plantas fuera de ruta
                palabras_excluidas = ['cartagena', 'ajover', 'bogota', 'altipal', 'cerete', 'serete', 'soberana']
                patron_exclusion = '|'.join(palabras_excluidas)
                
                # 2. Buscamos si la palabra prohibida está ÚNICAMENTE en la 'Sucursal'
                mascara_sucursal = df_limpio['Sucursal'].astype(str).str.contains(patron_exclusion, case=False, na=False)
                
                # 3. Nos quedamos con la flota filtrada: Todo lo que NO (~) tenga esas palabras en sucursal
                df_ruta_dino = df_limpio[~mascara_sucursal].copy()
                
                if not df_ruta_dino.empty:
                    total_visitas_ruta = len(df_ruta_dino)
                    horas_estimadas = total_visitas_ruta * 2
                    
                    c1, c2 = st.columns(2)
                    c1.metric("📍 Visitas en Ruta (Excluyendo foráneos)", total_visitas_ruta)
                    c2.metric("⏱️ Horas Técnicas Estimadas (2h x Eq)", f"{horas_estimadas} hrs")
                    
                    pdf_interno_bytes = generar_pdf_interno_dino(df_ruta_dino, horas_estimadas)
                    
                    st.download_button(
                        label="⚙️ Descargar Cronograma Interno (PDF)",
                        data=pdf_interno_bytes,
                        file_name="Cronograma_Interno_Dino.pdf",
                        mime="application/pdf"
                    )
                    
                    with st.expander("Ver lista de equipos incluidos en esta ruta"):
                        st.dataframe(df_ruta_dino[['Cliente', 'Equipo', 'Sucursal', 'Fecha_Visita']].sort_values(by=['Sucursal', 'Fecha_Visita']), hide_index=True)
                else:
                    st.warning("No se encontraron equipos para esta ruta después de aplicar las exclusiones.")

        except Exception as e:
            st.error(f"Error crítico al procesar el archivo. Detalles: {e}")
    else:
        st.info("👈 Sube el reporte de SAMM para gestionar los cronogramas.")


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
            while num > 40000:
                num = num / 10.0
            return num
        except:
            return None
            
    df_maestro['Horometro Actual'] = df_maestro['Horometro Actual'].apply(limpiar_horometro_base)
    df_maestro = df_maestro.dropna(subset=['Horometro Actual'])
    df_maestro['equipo_str'] = df_maestro['equipo'].astype(str).str.strip()
    
    archivo_tracker = "datos_samm/Control_Mantenimiento.xlsx"
    
    if os.path.exists(archivo_tracker):
        df_tracker_crudo = pd.read_excel(archivo_tracker)
        df_tracker = df_tracker_crudo.copy()
        
        if 'EQUIPO' in df_tracker.columns:
            df_tracker.rename(columns={'EQUIPO': 'Equipo'}, inplace=True)
        if 'Horometro de cambio deaceite' in df_tracker.columns:
            df_tracker.rename(columns={'Horometro de cambio deaceite': 'Ultimo_Mantenimiento'}, inplace=True)
            
        if 'ESTADO INSUMOS' not in df_tracker.columns:
            df_tracker['Estado_Insumos'] = "Al día"
        else:
            df_tracker.rename(columns={'ESTADO INSUMOS': 'Estado_Insumos'}, inplace=True)
            df_tracker['Estado_Insumos'] = df_tracker['Estado_Insumos'].fillna("Al día")
            
        if 'Horometro_Base_Ciclo' not in df_tracker.columns:
            df_tracker['Horometro_Base_Ciclo'] = df_tracker['Ultimo_Mantenimiento']
            
        df_tracker['Equipo'] = df_tracker['Equipo'].astype(str).str.strip()
        df_tracker = df_tracker[~df_tracker['Equipo'].str.upper().isin(['NAN', 'NONE', 'NA', ''])]
        df_maestro = df_maestro[~df_maestro['equipo_str'].str.upper().isin(['NAN', 'NONE', 'NA', ''])]
    else:
        st.warning("⚠️ No se encontró el archivo 'Control_Mantenimiento.xlsx' en 'datos_samm'.")
        st.stop()
        
    df_predictivo = pd.merge(
        df_tracker, 
        df_maestro[['equipo_str', 'Tercero', 'Sucursal', 'Modelo', 'Horometro Actual']], 
        left_on='Equipo', 
        right_on='equipo_str', 
        how='inner'
    )
    
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
        
        if faltan < -50:
            return "NIVEL MÁXIMO (Requiere Cambio TOTAL por Atraso)"
            
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
    
    st.subheader("🚨 Panel de Alertas y Trámites")
    
    df_alertas = df_predictivo[df_predictivo['Alerta'].str.contains('🔴|🟡')].sort_values(by='Horas_Faltantes')
    
    if not df_alertas.empty:
        columnas_ver = ['Equipo', 'Tercero', 'Horometro Actual', 'Horas_Faltantes', 'Tipo_Mantenimiento', 'Alerta', 'Estado_Insumos']
        st.dataframe(df_alertas[columnas_ver], hide_index=True)
        
        buffer_alertas = io.BytesIO()
        with pd.ExcelWriter(buffer_alertas, engine='xlsxwriter') as writer:
            df_alertas[columnas_ver].to_excel(writer, sheet_name='Solicitud_Filtros', index=False)
        
        st.download_button(
            label="📥 Descargar Reporte para Compras (Excel)",
            data=buffer_alertas.getvalue(),
            file_name="Solicitud_Insumos_Mantenimiento.xlsx",
            mime="application/vnd.ms-excel"
        )
        
        st.markdown("---")
        st.subheader("⚙️ Gestión de Estado del Mantenimiento")
        
        c1, c2, c3 = st.columns([2, 2, 1])
        
        with c1:
            equipo_revisado = st.selectbox("1. Selecciona el equipo:", df_alertas['Equipo'].tolist())
            
        with c2:
            nuevo_estado = st.selectbox(
                "2. Actualizar estado a:", 
                [
                    "Solicitado", 
                    "Pendiente por instalar", 
                    "✅ Cambio BÁSICO Realizado", 
                    "🚨 Cambio TOTAL Realizado (Reseteo)"
                ]
            )
            
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
                        st.success(f"¡Reseteo exitoso! El equipo {equipo_revisado} inicia un nuevo ciclo limpio desde {nuevo_horometro} hrs.")
                    else:
                        st.success(f"Mantenimiento básico registrado. El equipo {equipo_revisado} avanza en su ciclo normal.")
                    
                    from datetime import datetime
                    if 'Fecha de cambio aceite' in df_tracker.columns:
                        df_tracker.loc[df_tracker['Equipo'] == equipo_revisado, 'Fecha de cambio aceite'] = datetime.now().strftime("%Y-%m-%d")
                        
                else:
                    df_tracker.loc[df_tracker['Equipo'] == equipo_revisado, 'Estado_Insumos'] = nuevo_estado
                    st.success(f"Estado del equipo {equipo_revisado} actualizado a: {nuevo_estado}")
                
                df_salida = df_tracker.rename(columns={
                    'Equipo': 'EQUIPO',
                    'Ultimo_Mantenimiento': 'Horometro de cambio deaceite',
                    'Estado_Insumos': 'ESTADO INSUMOS'
                })
                df_salida.to_excel(archivo_tracker, index=False)
                st.rerun()
    else:
        st.success("✅ Toda la flota está en estado óptimo.")
        
    with st.expander("📊 Ver Flota Completa (Combustión) y Estados"):
        st.dataframe(df_predictivo[['Equipo', 'Tercero', 'Horometro Actual', 'Ultimo_Mantenimiento', 'Horometro_Base_Ciclo', 'Horas_Faltantes', 'Tipo_Mantenimiento', 'Estado_Insumos']], hide_index=True)

# =====================================================================
# 🚜 MÓDULO 3: DIRECTORIO DE FLOTA
# =====================================================================
elif menu_seleccionado == "🚜 3. Directorio de Flota":
    st.title("🚜 Directorio Global de Flota")
    st.markdown("---")
    
    df_dir = st.session_state['df_base_maestra'].copy()
    busqueda = st.text_input("🔍 Buscar por Número de Equipo, Cliente, Modelo o Sucursal:")
    
    if busqueda:
        df_dir = df_dir[df_dir.astype(str).apply(lambda x: x.str.contains(busqueda, case=False)).any(axis=1)]
        
    columnas_dir = ['equipo', 'Tercero', 'Sucursal', 'Modelo', 'Horometro Actual']
    
    if 'Link_Ficha' in df_dir.columns:
        columnas_dir.append('Link_Ficha')
        
    st.dataframe(df_dir[columnas_dir], use_container_width=True, hide_index=True)
    
    # --- BOTÓN DE DESCARGA DEL DIRECTORIO ---
    buffer_dir = io.BytesIO()
    with pd.ExcelWriter(buffer_dir, engine='xlsxwriter') as writer:
        df_dir[columnas_dir].to_excel(writer, sheet_name='Directorio', index=False)
        
    st.download_button(
        label="📥 Descargar Directorio de Pantalla (Excel)",
        data=buffer_dir.getvalue(),
        file_name="Directorio_Dinomontacargas.xlsx",
        mime="application/vnd.ms-excel"
    )