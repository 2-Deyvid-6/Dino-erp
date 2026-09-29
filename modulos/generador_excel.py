import pandas as pd
import re
from fpdf import FPDF
import os
from datetime import datetime
import calendar
# =====================================================================
# ESCUDO ANTIBALAS PARA CARACTERES RAROS (EVITA CRASHES DEL PDF)
# =====================================================================
def sanear_texto(txt):
    if pd.isna(txt) or txt is None: return ""
    txt = str(txt)
    # Reemplaza guiones y comillas de Word por caracteres estándar ASCII
    txt = txt.replace('–', '-').replace('—', '-') 
    txt = txt.replace('“', '"').replace('”', '"')
    txt = txt.replace('‘', "'").replace('’', "'")
    # Fuerza codificación compatible con la fuente Helvetica (Latin-1)
    return txt.encode('latin-1', 'ignore').decode('latin-1')

# =====================================================================
# MÓDULO 1: CONFIGURACIÓN BASE Y ESTILOS (El "Lienzo")
# =====================================================================
class ReportePDF(FPDF):
    # Interceptamos las funciones de dibujo para inyectar el escudo
    def cell(self, w, h=0, txt='', *args, **kwargs):
        txt_saneado = sanear_texto(txt)
        super().cell(w, h, txt_saneado, *args, **kwargs)

    def multi_cell(self, w, h, txt, *args, **kwargs):
        txt_saneado = sanear_texto(txt)
        super().multi_cell(w, h, txt_saneado, *args, **kwargs)

    def header(self):
        azul_dino = (0, 32, 96)
        
        ruta_logo = "logo.png"
        if os.path.exists(ruta_logo):
            self.image(ruta_logo, x=15, y=8, w=125)
        
        self.set_font("helvetica", "B", 16)
        self.set_text_color(*azul_dino)
        self.set_xy(10, 35)
        self.cell(0, 10, "CRONOGRAMA DE MANTENIMIENTO", ln=True, align='C')
        self.ln(2)
        
        self.set_font("helvetica", "I", 10)
        self.set_text_color(89, 89, 89)
        texto_intro = '"Con el fin de garantizar la continuidad operativa de sus equipos, compartimos la programación de servicio técnico correspondiente."'
        self.multi_cell(0, 5, texto_intro, align='C')
        self.ln(8)


# =====================================================================
# MÓDULO 2: UTILIDADES DE DATOS (El "Cerebro Limpiador")
# =====================================================================
def limpiar_equipo(texto_samm):
    texto = str(texto_samm).strip()
    # Busca ESTRICTAMENTE números dentro de corchetes [ ]
    match = re.search(r'\[\s*(\d+)\s*\]', texto)
    if match:
        return match.group(1).strip()
    # Si no tiene corchetes (ej. en la Base Maestra ya viene limpio), lo devuelve tal cual
    return texto

def nombre_mes(num_mes):
    meses = {1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL', 5: 'MAYO', 6: 'JUNIO',
             7: 'JULIO', 8: 'AGOSTO', 9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'}
    return meses.get(num_mes, "MES")

# =====================================================================
# MÓDULO 3: DIBUJADO DEL CALENDARIO (Vista Rápida)
# =====================================================================
def dibujar_calendario_dinamico(pdf, calendario, mes_str, num_dias):
    azul_dino = (0, 32, 96)
    naranja_dino = (255, 102, 0)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.set_text_color(*azul_dino)
    pdf.cell(0, 8, f"VISTA RÁPIDA DEL MES ({mes_str})", ln=True)
    
    x_start = pdf.get_x()
    y_line = pdf.get_y()
    pdf.set_draw_color(*naranja_dino)
    pdf.set_line_width(0.5)
    pdf.line(x_start, y_line, 200, y_line)
    pdf.ln(5)
    
    start_x = 10
    cell_w = 11.5 
    
    # Fila 1 (Días 1 al 16) - SIEMPRE FIJO
    y_days = pdf.get_y()
    y_eqs = y_days + 5
    max_y_row1 = y_eqs + 5

    for d in range(1, 17):
        pdf.set_xy(start_x + (d-1)*cell_w, y_days)
        pdf.set_font("helvetica", "B", 8)
        if calendario.get(d, []):
            pdf.set_text_color(*azul_dino)
        else:
            pdf.set_text_color(180, 180, 180)
        pdf.cell(cell_w, 5, str(d), align='C')
        
        pdf.set_xy(start_x + (d-1)*cell_w, y_eqs)
        pdf.set_font("helvetica", "B", 6)
        pdf.set_text_color(*naranja_dino)
        eq_text = ",".join(calendario.get(d, []))
        pdf.multi_cell(cell_w, 3.5, eq_text, align='C')
        if pdf.get_y() > max_y_row1: max_y_row1 = pdf.get_y()
    
    # Fila 2 (Días 17 hasta el FIN EXACTO DEL MES)
    pdf.set_y(max_y_row1 + 5)
    y_days = pdf.get_y()
    y_eqs = y_days + 5
    max_y_row2 = y_eqs + 5

    for d in range(17, num_dias + 1):
        pdf.set_xy(start_x + (d-17)*cell_w, y_days)
        pdf.set_font("helvetica", "B", 8)
        if calendario.get(d, []):
            pdf.set_text_color(*azul_dino)
        else:
            pdf.set_text_color(180, 180, 180)
        pdf.cell(cell_w, 5, str(d), align='C')
        
        pdf.set_xy(start_x + (d-17)*cell_w, y_eqs)
        pdf.set_font("helvetica", "B", 6)
        pdf.set_text_color(*naranja_dino)
        eq_text = ",".join(calendario.get(d, []))
        pdf.multi_cell(cell_w, 3.5, eq_text, align='C')
        if pdf.get_y() > max_y_row2: max_y_row2 = pdf.get_y()
            
    pdf.set_y(max_y_row2 + 10)

# =====================================================================
# MÓDULO 4: DIBUJADO DE TARJETAS (Vista Detallada)
# =====================================================================
def dibujar_tarjetas_equipos(pdf, df_mes, mes_str):
    azul_dino = (0, 32, 96)
    naranja_dino = (255, 102, 0)
    gris_texto = (89, 89, 89)

    pdf.set_font("helvetica", "B", 11)
    pdf.set_text_color(*azul_dino)
    pdf.cell(0, 8, f"VISTA DETALLADA DEL MES ({mes_str})", ln=True)
    
    x_start = pdf.get_x()
    y_line = pdf.get_y()
    pdf.set_draw_color(*naranja_dino)
    pdf.set_line_width(0.5)
    pdf.line(x_start, y_line, 200, y_line)
    pdf.ln(6)

    equipos_agrupados = df_mes.groupby('EquipoLimpio')

    for equipo_limpio, datos in equipos_agrupados:
        sucursal = str(datos['Sucursal'].iloc[0]).strip()
        ciudad = str(datos['Ciudad'].iloc[0]).strip() if 'Ciudad' in datos.columns else ""
        
        if sucursal.lower() in ["sin crear", "nan", ""]: sucursal = ""
        if ciudad.lower() in ["nan", ""]: ciudad = ""
        
        texto_ubicacion = sucursal
        if ciudad and ciudad.lower() not in sucursal.lower():
            texto_ubicacion += f" - {ciudad}" if sucursal else ciudad
            
        fechas_visitas = datos['Fecha_Visita'].sort_values().tolist()
        num_visitas = len(fechas_visitas)
        
        if pdf.get_y() + 38 > 280:
            pdf.add_page()

        block_y = pdf.get_y()
        block_h = 36
        
        # 1. Tarjeta Base
        pdf.set_draw_color(220, 220, 220)
        pdf.set_line_width(0.3)
        pdf.rect(x=10, y=block_y, w=190, h=block_h, style='D', round_corners=True, corner_radius=3)
        
        # 2. Línea Naranja lateral
        pdf.set_draw_color(*naranja_dino)
        pdf.set_line_width(1.5)
        pdf.line(13, block_y + 5, 13, block_y + 15)

        # 3. TÍTULO
        pdf.set_xy(16, block_y + 4)
        pdf.set_font("helvetica", "B", 13)
        pdf.set_text_color(*azul_dino)
        pdf.cell(0, 6, f"Equipo: {equipo_limpio}", ln=True)
        
        # 4. LÍNEA DELGADA
        pdf.set_draw_color(210, 210, 210)
        pdf.set_line_width(0.2)
        pdf.line(16, block_y + 11.5, 195, block_y + 11.5)

        # 5. UBICACIÓN TÉCNICA
        pdf.set_xy(16, block_y + 13)
        pdf.set_font("helvetica", "B", 9)
        pdf.set_text_color(*naranja_dino)
        if texto_ubicacion:
            pdf.cell(0, 5, f"Ubicación técnica: {texto_ubicacion}", ln=True)
        else:
            pdf.cell(0, 5, "Ubicación técnica:", ln=True)
            
        # 6. TABLAS DE FECHAS
        table_y = block_y + 20
        pdf.set_line_width(0.2)
        pdf.set_draw_color(191, 191, 191)
        
        box_width = min(85, 178 / num_visitas) if num_visitas > 0 else 85
        total_boxes_width = box_width * num_visitas
        start_x_centered = (210 - total_boxes_width) / 2
        
        pdf.rect(x=start_x_centered, y=table_y, w=total_boxes_width, h=11, style='D', round_corners=True, corner_radius=2)
        pdf.line(start_x_centered, table_y + 5, start_x_centered + total_boxes_width, table_y + 5)
        
        for i in range(1, num_visitas):
            line_x = start_x_centered + (i * box_width)
            pdf.line(line_x, table_y, line_x, table_y + 11)
            
        pdf.set_font("helvetica", "B", 8)
        pdf.set_text_color(*gris_texto)
        for i in range(num_visitas):
            pdf.set_xy(start_x_centered + (i * box_width), table_y)
            pdf.cell(box_width, 5, f"VISITA TÉCNICA {i+1}", border=0, align="C")
            
        pdf.set_font("helvetica", "B", 9)
        pdf.set_text_color(*azul_dino)
        for i, fecha in enumerate(fechas_visitas):
            pdf.set_xy(start_x_centered + (i * box_width), table_y + 5)
            fecha_limpia = str(fecha).split(" ")[0].strip() if pd.notna(fecha) else "S/F"
            pdf.cell(box_width, 6, fecha_limpia, border=0, align="C")
            
        pdf.set_y(block_y + block_h + 4)

# =====================================================================
# MÓDULO 4.5: TARJETA HERO DEL CLIENTE
# =====================================================================
def dibujar_tarjeta_cliente(pdf, nombre_cliente, num_equipos, ciudad, nit):
    azul_dino = (0, 32, 96)
    naranja_dino = (255, 102, 0)
    
    pdf.ln(5) 
    y_start = pdf.get_y()
    
    pdf.set_draw_color(*azul_dino)
    pdf.set_line_width(0.4)
    pdf.rect(x=10, y=y_start, w=190, h=22, style='D', round_corners=True, corner_radius=3)
    
    pdf.set_xy(10, y_start + 4)
    pdf.set_font("helvetica", "B", 14)
    pdf.set_text_color(*azul_dino)
    pdf.cell(190, 6, str(nombre_cliente).upper(), align='C')
    
    pdf.set_draw_color(220, 220, 220)
    pdf.set_line_width(0.3)
    pdf.line(15, y_start + 12, 195, y_start + 12)
    
    pdf.set_xy(10, y_start + 14)
    pdf.set_font("helvetica", "B", 9)
    pdf.set_text_color(*naranja_dino)
    
    texto_inferior = f"TOTAL EQUIPOS: {num_equipos}   |   NIT: {nit}"
    pdf.cell(190, 5, texto_inferior, align='C')
    
    pdf.set_y(y_start + 30)

# =====================================================================
# MÓDULO 4.8: TARJETAS DE PIE DE PÁGINA (ELÁSTICAS Y SEGURAS)
# =====================================================================
def dibujar_footer_informativo(pdf):
    azul_dino = (0, 32, 96)
    naranja_dino = (255, 102, 0)
    gris_texto = (89, 89, 89)

    if pdf.get_y() > 215:
        pdf.add_page()
        
    pdf.ln(8)

    pdf.set_font("helvetica", "B", 10)
    pdf.set_text_color(*azul_dino)
    pdf.cell(0, 6, "Personal Técnico Asignado:", ln=True)
    
    y_start = pdf.get_y()
    tecnicos = "ARIEL ORTEGA  |  GERMAN LECHUGA  |  ISAAC MARIOTA  |  EDGAR SOLANO  | JESUS AVENDAÑO  | SANTIAGO PINTO |  MARTIN SUAREZ | YOINER HURTADO"
    
    pdf.set_xy(12, y_start + 3.5)
    pdf.set_font("helvetica", "B", 8)
    pdf.set_text_color(*gris_texto)
    pdf.multi_cell(186, 5, tecnicos, align='C')
    
    y_end = pdf.get_y()
    box_height = (y_end - y_start) + 3.5
    
    pdf.set_draw_color(220, 220, 220)
    pdf.set_line_width(0.3)
    pdf.rect(x=10, y=y_start, w=190, h=box_height, style='D', round_corners=True, corner_radius=2)
    
    pdf.set_y(y_end + 8)

    if pdf.get_y() > 260:
        pdf.add_page()

    y_start = pdf.get_y()
    pdf.set_draw_color(*azul_dino)
    pdf.set_line_width(0.4)
    pdf.rect(x=10, y=y_start, w=190, h=18, style='D', round_corners=True, corner_radius=3)
    
    auto_pb = pdf.auto_page_break
    pb_margin = pdf.b_margin
    pdf.set_auto_page_break(False)
    
    pdf.set_xy(10, y_start + 4)
    pdf.set_font("helvetica", "B", 11)
    pdf.set_text_color(*azul_dino)
    pdf.cell(190, 5, "HORARIO DE ATENCIÓN DE NUESTRO EQUIPO TÉCNICO DINO: LUNES A VIERNES 7:40 A 4:50", align='C')
    
    pdf.set_xy(10, y_start + 10)
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(190, 5, "SÁBADO DE 8:10 A 11:50", align='C')
    
    pdf.set_auto_page_break(auto_pb, pb_margin)
    pdf.set_y(y_start + 18)

# =====================================================================
# MÓDULO 5: DIRECTOR DE ORQUESTA (CANDADO DE MES ÚNICO)
# =====================================================================
def generar_pdf_cliente(df_cliente, nombre_cliente, texto_novedad="", lista_combustion=None):
    if lista_combustion is None: lista_combustion = []
    pdf = ReportePDF(orientation='P', unit='mm', format='A4')
    
    # Pre-procesamiento de datos
    df_cliente_clean = df_cliente.copy()
    df_cliente_clean['EquipoLimpio'] = df_cliente_clean['Equipo'].apply(limpiar_equipo)
    
    def extraer_mes_año(fecha_str):
        try:
            f = str(fecha_str).split(" ")[0]
            obj = datetime.strptime(f, "%d/%m/%Y")
            return (obj.year, obj.month)
        except:
            return (9999, 99) 
            
    df_cliente_clean['Mes_Clave'] = df_cliente_clean['Fecha_Visita'].apply(extraer_mes_año)
    
    # --- CANDADO ESTRICTO DE MES ÚNICO ---
    if not df_cliente_clean.empty:
        meses_validos = df_cliente_clean[df_cliente_clean['Mes_Clave'] != (9999, 99)]
        if not meses_validos.empty:
            mes_principal = meses_validos['Mes_Clave'].mode()[0]
            df_cliente_clean = df_cliente_clean[df_cliente_clean['Mes_Clave'] == mes_principal]
            
    # --- PURGA DE VISITAS DUPLICADAS ---
    # Si un equipo tiene múltiples visitas el MISMO día, se fusionan en una sola para el PDF
    df_cliente_clean = df_cliente_clean.drop_duplicates(subset=['EquipoLimpio', 'Fecha_Visita'])

    num_equipos = df_cliente_clean['EquipoLimpio'].nunique() 
    nit_cliente = str(df_cliente_clean['NIT'].iloc[0]) if 'NIT' in df_cliente_clean.columns else "S/D"
    ciudad_cliente = str(df_cliente_clean['Ciudad_Global'].iloc[0]) if 'Ciudad_Global' in df_cliente_clean.columns else "S/D"

    meses_presentes = sorted(df_cliente_clean['Mes_Clave'].unique())

    # De aquí en adelante el código sigue intacto
    for i, mes_clave in enumerate(meses_presentes):
        anio, mes_num = mes_clave
        if anio == 9999: continue 
        
        num_dias = calendar.monthrange(anio, mes_num)[1]
        
        mes_str = f"{nombre_mes(mes_num)} {anio}"
        df_mes_actual = df_cliente_clean[df_cliente_clean['Mes_Clave'] == mes_clave]
        
        # 3. Armar el calendario de la vista rápida
        calendario_mes = {d: [] for d in range(1, num_dias + 1)}
        for _, row in df_mes_actual.iterrows():
            try:
                f_str = str(row['Fecha_Visita']).split(" ")[0]
                dia = datetime.strptime(f_str, "%d/%m/%Y").day
                eq = row['EquipoLimpio']
                
                if dia <= num_dias:
                    if eq not in calendario_mes[dia]: calendario_mes[dia].append(eq)
            except:
                pass
                
        pdf.add_page()
        if i == 0:
            dibujar_tarjeta_cliente(pdf, nombre_cliente, num_equipos, ciudad_cliente, nit_cliente)
            
        dibujar_calendario_dinamico(pdf, calendario_mes, mes_str, num_dias)
        
        # Dibuja las tarjetas manteniendo el nombre original de la sucursal intacto
        dibujar_tarjetas_equipos(pdf, df_mes_actual, mes_str)


    # =========================================================
    # NUEVAS TARJETAS (DISEÑO UNIFICADO Y ELÁSTICO)
    # =========================================================
    pdf.ln(5)
    
    equipos_cliente = df_cliente_clean['EquipoLimpio'].astype(str).str.strip().unique()
    tiene_combustion = any(eq in lista_combustion for eq in equipos_cliente)
    tiene_electrico = any(eq not in lista_combustion for eq in equipos_cliente)

    rutina_combustion = (
        "- Cambio de aceite de motor, filtros de aceite, aire y combustible.\n"
        "- Revision de sistema de refrigeracion, correas y tension.\n"
        "- Chequeo de nivel de aceite hidraulico y de transmision.\n"
        "- Revision de frenos, sistema electrico, luces y alarmas.\n"
        "- Lubricacion y engrase general de mastil, cadenas y rodamientos."
    )

    rutina_electrica = (
        "- Revision profunda de baterias (electrolito, bornes, limpieza).\n"
        "- Chequeo de contactores, tarjeta de control y arneses.\n"
        "- Revision de motores de traccion y bombeo (escobillas si aplica).\n"
        "- Chequeo de nivel de aceite hidraulico y control de fugas.\n"
        "- Revision de frenos, luces, alarmas y engrase de mastil/cadenas."
    )

    if tiene_combustion:
        if pdf.get_y() > 230: pdf.add_page()
        
        pdf.set_font("helvetica", "B", 10)
        pdf.set_text_color(0, 32, 96) 
        pdf.cell(0, 6, "RUTINA PREVENTIVA - EQUIPOS A COMBUSTION:", ln=True)
        
        y_start = pdf.get_y()
        pdf.set_xy(12, y_start + 3)
        pdf.set_font("helvetica", "", 9)
        pdf.set_text_color(89, 89, 89)
        pdf.multi_cell(186, 4.5, rutina_combustion)
        
        y_end = pdf.get_y()
        box_height = (y_end - y_start) + 3 # Calcula altura dinámica
        
        pdf.set_draw_color(220, 220, 220)
        pdf.set_line_width(0.3)
        pdf.rect(x=10, y=y_start, w=190, h=box_height, style='D', round_corners=True, corner_radius=2)
        pdf.set_y(y_end + 8)

    if tiene_electrico:
        if pdf.get_y() > 230: pdf.add_page()
        
        pdf.set_font("helvetica", "B", 10)
        pdf.set_text_color(0, 32, 96)
        pdf.cell(0, 6, "RUTINA PREVENTIVA - EQUIPOS ELECTRICOS:", ln=True)
        
        y_start = pdf.get_y()
        pdf.set_xy(12, y_start + 3)
        pdf.set_font("helvetica", "", 9)
        pdf.set_text_color(89, 89, 89)
        pdf.multi_cell(186, 4.5, rutina_electrica)
        
        y_end = pdf.get_y()
        box_height = (y_end - y_start) + 3
        
        pdf.set_draw_color(220, 220, 220)
        pdf.set_line_width(0.3)
        pdf.rect(x=10, y=y_start, w=190, h=box_height, style='D', round_corners=True, corner_radius=2)
        pdf.set_y(y_end + 8)

    if texto_novedad and texto_novedad.strip() != "":
        if pdf.get_y() > 240: pdf.add_page()
        
        pdf.set_font("helvetica", "B", 10)
        pdf.set_text_color(255, 102, 0) 
        pdf.cell(0, 6, "¡UN MENSAJE DE NOSOTROS, PARA USTEDES!:", ln=True)
        
        y_start = pdf.get_y()
        pdf.set_xy(12, y_start + 3)
        pdf.set_font("helvetica", "I", 9) 
        pdf.set_text_color(89, 89, 89)
        pdf.multi_cell(186, 4.5, texto_novedad)
        
        y_end = pdf.get_y()
        box_height = (y_end - y_start) + 3 # Altura dinámica para novedades largas
        
        pdf.set_draw_color(220, 220, 220)
        pdf.set_line_width(0.3)
        pdf.rect(x=10, y=y_start, w=190, h=box_height, style='D', round_corners=True, corner_radius=2)
        pdf.set_y(y_end + 8)

    # 3. Insertamos el footer
    dibujar_footer_informativo(pdf)

    return bytes(pdf.output())

# =====================================================================
# MÓDULO 6: CRONOGRAMA INTERNO LOGÍSTICO (RUTA DINO)
# =====================================================================

def generar_pdf_interno_dino(df_ruta, horas_estimadas):
    # 1. Preparación y Limpieza de Fechas
    df = df_ruta.copy()
    
    df['Fecha_DT'] = pd.to_datetime(df['Fecha_Visita'], format='%d/%m/%Y', errors='coerce')
    df['Fecha_DT'] = df['Fecha_DT'].fillna(pd.to_datetime(df['Fecha_Visita'], errors='coerce'))
    df = df.dropna(subset=['Fecha_DT']).sort_values('Fecha_DT')
    
    df['Año_ISO'] = df['Fecha_DT'].dt.isocalendar().year
    df['Semana_ISO'] = df['Fecha_DT'].dt.isocalendar().week
    
    dias_semana_es = {0: 'LUNES', 1: 'MARTES', 2: 'MIÉRCOLES', 3: 'JUEVES', 4: 'VIERNES', 5: 'SÁBADO', 6: 'DOMINGO'}
    
    # 2. Inicialización del PDF usando nuestra clase blindada
    pdf = ReportePDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    azul_dino = (0, 32, 96)
    naranja_dino = (255, 102, 0)
    gris_borde = (200, 200, 200)
    
    # Encabezado (Empujado a Y=65 para NO chocar con el membrete fijo de ReportePDF)
    pdf.set_y(65)
    pdf.set_font("helvetica", "B", 14)
    pdf.set_text_color(*azul_dino)
    pdf.cell(0, 8, "RUTERO LOGÍSTICO (ZONA URBANA)", ln=True, align="C")
    
    pdf.set_font("helvetica", "I", 10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 6, f"Total Visitas Programadas: {len(df)} | Carga Técnica Estimada: {horas_estimadas} hrs", ln=True, align="C")
    pdf.ln(6)

    # 3. Motor de Construcción (Semanas -> Días -> Tablas)
    grupos_semana = df.groupby(['Año_ISO', 'Semana_ISO'])
    contador_semana = 1
    
    for (anio, semana_iso), df_sem in grupos_semana:
        fecha_min = df_sem['Fecha_DT'].min()
        fecha_max = df_sem['Fecha_DT'].max()
        
        # Evitar que el título de la semana quede cortado al final de la página
        if pdf.get_y() > 240:
            pdf.add_page()
            pdf.set_y(65)

        # Título de Semana
        pdf.set_font("helvetica", "B", 12)
        pdf.set_fill_color(*azul_dino)
        pdf.set_text_color(255, 255, 255)
        rango_fechas = f"Del {fecha_min.strftime('%d/%m/%Y')} al {fecha_max.strftime('%d/%m/%Y')}"
        pdf.cell(0, 8, f"  SEMANA {contador_semana}  -  {rango_fechas}", ln=True, fill=True)
        pdf.ln(5)
        
        grupos_dia = df_sem.groupby('Fecha_DT')
        for fecha_dia, df_dia in grupos_dia:
            nombre_dia = dias_semana_es[fecha_dia.dayofweek]
            fecha_str = fecha_dia.strftime('%d/%m/%Y')
            titulo_dia = f"📅 {nombre_dia} {fecha_str} (Equipos: {len(df_dia)})"
            
            # 🔥 MOTOR DE PAGINACIÓN INTELIGENTE (CHUNKING PARA TABLAS)
            df_restante = df_dia.copy()
            es_primer_bloque = True
            
            while not df_restante.empty:
                if pdf.get_y() > 265:
                    pdf.add_page()
                    pdf.set_y(65)

                block_y = pdf.get_y()
                espacio_disponible = 275 - block_y
                
                # Altura dinámica: si es el inicio del día requiere 14mm (Título + Cabeceras), si no, solo 7mm (Cabeceras)
                alto_encabezado = 14 if es_primer_bloque else 7
                espacio_para_filas = espacio_disponible - alto_encabezado
                filas_que_caben = int(espacio_para_filas // 7)
                
                if filas_que_caben < 1:
                    pdf.add_page()
                    pdf.set_y(65)
                    continue

                chunk = df_restante.iloc[:filas_que_caben]
                df_restante = df_restante.iloc[filas_que_caben:]
                
                altura_tarjeta = alto_encabezado + (len(chunk) * 7)
                
                # 1. Dibujar la cápsula contenedora perfecta (Tabla principal)
                pdf.set_draw_color(*gris_borde)
                pdf.set_line_width(0.3)
                try:
                    pdf.rect(x=10, y=block_y, w=190, h=altura_tarjeta, style='D', round_corners=True, corner_radius=2)
                except:
                    pdf.rect(x=10, y=block_y, w=190, h=altura_tarjeta, style='D')

                y_actual = block_y

                # 2. Título del Día (Solo se pinta la primera vez que se abre la tabla del día)
                if es_primer_bloque:
                    pdf.set_xy(10, y_actual)
                    pdf.set_font("helvetica", "B", 10)
                    pdf.set_text_color(*naranja_dino)
                    pdf.cell(190, 7, f"  {titulo_dia}", border='B', align='L')
                    y_actual += 7

                # 3. Cabeceras de Columnas
                pdf.set_xy(10, y_actual)
                pdf.set_fill_color(245, 245, 245)
                pdf.set_font("helvetica", "B", 8)
                pdf.set_text_color(*azul_dino)
                col_w = [22, 90, 78] # Anchos de columna [Equipo, Cliente, Sucursal]

                pdf.cell(col_w[0], 7, "EQUIPO", border='R', fill=True, align='C')
                pdf.cell(col_w[1], 7, "CLIENTE", border='R', fill=True, align='C')
                pdf.cell(col_w[2], 7, "SUCURSAL", border=0, fill=True, align='C')
                y_actual += 7
                pdf.line(10, y_actual, 200, y_actual)

                # 4. Filas
                pdf.set_font("helvetica", "", 8)
                pdf.set_text_color(60, 60, 60)
                
                for idx_row, (_, fila) in enumerate(chunk.iterrows()):
                    pdf.set_xy(10, y_actual)
                    
                    equipo = str(fila.get('Equipo', 'N/A')).strip()
                    cliente = str(fila.get('Cliente', 'N/A')).strip()
                    sucursal = str(fila.get('Sucursal', 'N/A')).strip()
                    
                    # Truncamiento inteligente para que no desborde las celdas
                    if len(cliente) > 55: cliente = cliente[:52] + "..."
                    if len(sucursal) > 48: sucursal = sucursal[:45] + "..."
                    
                    pdf.cell(col_w[0], 7, equipo, border='R', align='C')
                    pdf.cell(col_w[1], 7, f"  {cliente}", border='R', align='L')
                    pdf.cell(col_w[2], 7, f"  {sucursal}", border=0, align='L')
                    
                    y_actual += 7
                    # Dibujar línea horizontal si no es la última fila del bloque
                    if idx_row < len(chunk) - 1:
                        pdf.line(10, y_actual, 200, y_actual)

                pdf.set_y(block_y + altura_tarjeta + 6) # Margen inferior de la tabla
                es_primer_bloque = False
                
        pdf.ln(4)
        contador_semana += 1
        
    # 4. Exportación blindada
    resultado_pdf = pdf.output(dest='S')
    if isinstance(resultado_pdf, str):
        return resultado_pdf.encode('latin-1', 'ignore')
    else:
        return bytes(resultado_pdf)
