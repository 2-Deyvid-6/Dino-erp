import pandas as pd
import re
from fpdf import FPDF
import os
from datetime import datetime
from collections import defaultdict

# =====================================================================
# MÓDULO 1: CONFIGURACIÓN BASE Y ESTILOS (El "Lienzo")
# =====================================================================
class ReportePDF(FPDF):
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
    match = re.search(r'\[\s*(.*?)\s*\]', texto)
    if match:
        return match.group(1).strip()
    numeros = re.findall(r'\d+', texto)
    if numeros:
        return numeros[-1]
    return texto

def nombre_mes(num_mes):
    meses = {1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL', 5: 'MAYO', 6: 'JUNIO',
             7: 'JULIO', 8: 'AGOSTO', 9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'}
    return meses.get(num_mes, "MES")

# =====================================================================
# MÓDULO 3: DIBUJADO DEL CALENDARIO (Vista Rápida)
# =====================================================================
def dibujar_calendario_dinamico(pdf, calendario, mes_str):
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
    
    # Fila 1 (Días 1 al 16)
    y_days = pdf.get_y()
    y_eqs = y_days + 5
    max_y_row1 = y_eqs + 5

    for d in range(1, 17):
        pdf.set_xy(start_x + (d-1)*cell_w, y_days)
        pdf.set_font("helvetica", "B", 8)
        if calendario[d]:
            pdf.set_text_color(*azul_dino)
        else:
            pdf.set_text_color(180, 180, 180)
        pdf.cell(cell_w, 5, str(d), align='C')
        
        pdf.set_xy(start_x + (d-1)*cell_w, y_eqs)
        pdf.set_font("helvetica", "B", 6)
        pdf.set_text_color(*naranja_dino)
        eq_text = ",".join(calendario[d])
        pdf.multi_cell(cell_w, 3.5, eq_text, align='C')
        if pdf.get_y() > max_y_row1: max_y_row1 = pdf.get_y()
    
    # Fila 2 (Días 17 al 31)
    pdf.set_y(max_y_row1 + 5)
    y_days = pdf.get_y()
    y_eqs = y_days + 5
    max_y_row2 = y_eqs + 5

    for d in range(17, 32):
        pdf.set_xy(start_x + (d-17)*cell_w, y_days)
        pdf.set_font("helvetica", "B", 8)
        if calendario[d]:
            pdf.set_text_color(*azul_dino)
        else:
            pdf.set_text_color(180, 180, 180)
        pdf.cell(cell_w, 5, str(d), align='C')
        
        pdf.set_xy(start_x + (d-17)*cell_w, y_eqs)
        pdf.set_font("helvetica", "B", 6)
        pdf.set_text_color(*naranja_dino)
        eq_text = ",".join(calendario[d])
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
    
    # AJUSTE: Se eliminó la variable "ciudad" para evitar conflictos multi-sede
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

    # 1. AUMENTAMOS EL MARGEN DE SEGURIDAD GLOBAL (De 230 a 215)
    if pdf.get_y() > 215:
        pdf.add_page()
        
    pdf.ln(8)

    # Título Flotante
    pdf.set_font("helvetica", "B", 10)
    pdf.set_text_color(*azul_dino)
    pdf.cell(0, 6, "Personal Técnico Asignado:", ln=True)
    
    # 2. Caja Adaptativa de Técnicos
    y_start = pdf.get_y()
    tecnicos = "ARIEL ORTEGA  |  GERMAN LECHUGA  |  ISAAC MARIOTA  |  EDGAR SOLANO  | JESUS AVENDAÑO  | SANTIAGO PINTO |  MARTIN SUAREZ | YOINER HURTADO"
    
    pdf.set_xy(12, y_start + 3.5)
    pdf.set_font("helvetica", "B", 8)
    pdf.set_text_color(*gris_texto)
    pdf.multi_cell(186, 5, tecnicos, align='C')
    
    # Cálculo exacto de la altura del recuadro
    y_end = pdf.get_y()
    box_height = (y_end - y_start) + 3.5
    
    pdf.set_draw_color(220, 220, 220)
    pdf.set_line_width(0.3)
    pdf.rect(x=10, y=y_start, w=190, h=box_height, style='D', round_corners=True, corner_radius=2)
    
    pdf.set_y(y_end + 8)

    # 3. SEGUNDA VALIDACIÓN DE SEGURIDAD (Por si el bloque anterior creció mucho)
    if pdf.get_y() > 260:
        pdf.add_page()

    # 4. Caja de Horario
    y_start = pdf.get_y()
    pdf.set_draw_color(*azul_dino)
    pdf.set_line_width(0.4)
    pdf.rect(x=10, y=y_start, w=190, h=18, style='D', round_corners=True, corner_radius=3)
    
    # Apagamos temporalmente el salto automático para forzar el dibujo dentro de la caja
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
    
    # Reactivamos la configuración normal del PDF
    pdf.set_auto_page_break(auto_pb, pb_margin)
    pdf.set_y(y_start + 18)

# =====================================================================
# MÓDULO 5: DIRECTOR DE ORQUESTA (NUEVA LÓGICA MULTI-MES)
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
    
    num_equipos = df_cliente_clean['EquipoLimpio'].nunique() 
    nit_cliente = str(df_cliente_clean['NIT'].iloc[0]) if 'NIT' in df_cliente_clean.columns else "S/D"
    ciudad_cliente = str(df_cliente_clean['Ciudad_Global'].iloc[0]) if 'Ciudad_Global' in df_cliente_clean.columns else "S/D"

    meses_presentes = sorted(df_cliente_clean['Mes_Clave'].unique())

    for i, mes_clave in enumerate(meses_presentes):
        anio, mes_num = mes_clave
        if anio == 9999: continue 
        
        mes_str = f"{nombre_mes(mes_num)} {anio}"
        df_mes_actual = df_cliente_clean[df_cliente_clean['Mes_Clave'] == mes_clave]
        
        calendario_mes = {d: [] for d in range(1, 32)}
        for _, row in df_mes_actual.iterrows():
            try:
                f_str = str(row['Fecha_Visita']).split(" ")[0]
                dia = datetime.strptime(f_str, "%d/%m/%Y").day
                eq = row['EquipoLimpio']
                if eq not in calendario_mes[dia]: calendario_mes[dia].append(eq)
            except:
                pass
                
        pdf.add_page()
        if i == 0:
            dibujar_tarjeta_cliente(pdf, nombre_cliente, num_equipos, ciudad_cliente, nit_cliente)
            
        dibujar_calendario_dinamico(pdf, calendario_mes, mes_str)
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
        pdf.cell(0, 6, "NOVEDADES Y OBSERVACIONES DEL MES:", ln=True)
        
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
def generar_pdf_interno_dino(df_ruta, horas_totales):
    pdf = ReportePDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    azul_dino = (0, 32, 96)
    naranja_dino = (255, 102, 0)
    
    # --- TARJETA HERO (MANTENIMIENTO DINO) ---
    pdf.ln(5)
    pdf.set_draw_color(*azul_dino)
    pdf.set_line_width(0.4)
    pdf.rect(x=10, y=pdf.get_y(), w=190, h=22, style='D', round_corners=True, corner_radius=3)
    
    pdf.set_xy(10, pdf.get_y() + 4)
    pdf.set_font("helvetica", "B", 14)
    pdf.set_text_color(*azul_dino)
    pdf.cell(190, 6, "MANTENIMIENTO DINO - RUTA METROPOLITANA", align='C')
    
    pdf.set_xy(10, pdf.get_y() + 8)
    pdf.set_font("helvetica", "B", 10)
    pdf.set_text_color(*naranja_dino)
    pdf.cell(190, 5, f"TOTAL VISITAS: {len(df_ruta)}   |   TIEMPO ESTIMADO: {horas_totales} HORAS TÉCNICAS", align='C')
    pdf.ln(15)
    
    # --- AGRUPACIÓN POR CLIENTE ---
    clientes_agrupados = df_ruta.groupby('Cliente')
    
    for cliente, datos in clientes_agrupados:
        if pdf.get_y() > 250:
            pdf.add_page()
            
        pdf.set_font("helvetica", "B", 11)
        pdf.set_text_color(*azul_dino)
        pdf.cell(0, 8, f"CLIENTE: {str(cliente).upper()}", ln=True)
        
        pdf.set_fill_color(240, 240, 240)
        pdf.set_font("helvetica", "B", 9)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(30, 6, "FECHA", border=1, fill=True, align='C')
        pdf.cell(35, 6, "EQUIPO", border=1, fill=True, align='C')
        pdf.cell(125, 6, "UBICACIÓN FÍSICA (SUCURSAL)", border=1, fill=True, align='C')
        pdf.ln()
        
        pdf.set_font("helvetica", "", 8)
        for _, row in datos.sort_values('Fecha_Visita').iterrows():
            fecha = str(row['Fecha_Visita']).split(" ")[0] if pd.notna(row['Fecha_Visita']) else "S/F"
            eq = str(row['Equipo'])
            suc = str(row['Sucursal']).strip()
            if len(suc) > 75: suc = suc[:72] + "..." 
            
            pdf.cell(30, 6, fecha, border=1, align='C')
            pdf.cell(35, 6, eq, border=1, align='C')
            pdf.cell(125, 6, suc, border=1, align='L')
            pdf.ln()
            
        pdf.ln(6) 
        
    return bytes(pdf.output())