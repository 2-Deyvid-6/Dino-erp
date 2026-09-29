import pandas as pd
import numpy as np

def limpiar_reporte_samm(ruta_archivo):
    df_raw = pd.read_excel(ruta_archivo, header=None)
    
    # 1. Variables de memoria
    cliente_actual, nit_actual, ciudad_global_actual, fecha_actual = "DESCONOCIDO", "NO REGISTRA", "NO REGISTRA", "DESCONOCIDA"
    header_idx = -1
    
    # 2. ESCANEO RÁPIDO: Solo buscamos en las primeras 100 filas (para encontrar metadatos y cabeceras)
    for idx, row in df_raw.head(100).iterrows():
        valores = [str(x).strip() for x in row.values if pd.notna(x)]
        if not valores: continue
        
        # Atrapamos Metadatos
        if "CLIENTE:" in valores: cliente_actual = valores[valores.index("CLIENTE:") + 1] if len(valores) > valores.index("CLIENTE:") + 1 else cliente_actual
        if "NIT:" in valores: nit_actual = valores[valores.index("NIT:") + 1] if len(valores) > valores.index("NIT:") + 1 else nit_actual
        if "CIUDAD:" in valores: ciudad_global_actual = valores[valores.index("CIUDAD:") + 1] if len(valores) > valores.index("CIUDAD:") + 1 else ciudad_global_actual
        
        for val in valores:
            if val.startswith("Fecha Visita:"): fecha_actual = val.replace("Fecha Visita:", "").strip()
            
        # Detectamos la fila donde empiezan realmente las columnas de datos
        if "Equipo" in valores and ("Sucursal" in valores or "Visita" in valores):
            header_idx = idx
            break # DETENEMOS el bucle aquí. Cero iteraciones para el resto de los datos.

    if header_idx == -1:
        # Falla segura si el documento no tiene el formato esperado
        df_vacio = pd.DataFrame(columns=['Cliente', 'NIT', 'Ciudad_Global', 'Fecha_Visita', 'Equipo', 'Mantenimiento', 'OT', 'Estado', 'Sucursal', 'Ciudad', 'Color_Semantico'])
        return df_vacio
        
    # 3. VECTORIZACIÓN MASIVA (Aquí ocurre la magia de optimización)
    # Cortamos el DataFrame desde la fila de cabeceras hacia abajo
    df_data = df_raw.iloc[header_idx + 1:].copy()
    df_data.columns = df_raw.iloc[header_idx].astype(str).str.strip()
    
    # Renombrar dinámicamente columnas clave si existen (Manejo de mayúsculas/minúsculas)
    cols_map = {c: c.capitalize() for c in df_data.columns if str(c).lower() in ['equipo', 'visita', 'ot', 'estado', 'sucursal', 'ciudad']}
    df_data.rename(columns=cols_map, inplace=True)
    if 'Ot' in df_data.columns: df_data.rename(columns={'Ot': 'OT'}, inplace=True)
    
    # 4. LIMPIEZA DE BASURA SIN BUCLES (Filtros booleanos)
    df_data = df_data.dropna(subset=['Equipo'])
    palabras_basura = ('EQUIPO', 'CLIENTE:', 'NIT:', 'CIUDAD', 'DIRECCION', 'CONTRATO', 'FECHA', 'NAN')
    df_data = df_data[~df_data['Equipo'].astype(str).str.upper().str.startswith(palabras_basura)]
    
    # 5. RELLENO DE DATOS Y NORMALIZACIÓN (Vectorizado)
    for col, default_val in [('OT', 'sin crear'), ('Estado', 'Programada'), ('Sucursal', ''), ('Ciudad', ''), ('Visita', '')]:
        if col not in df_data.columns: df_data[col] = default_val
        
    df_data['OT'] = df_data['OT'].astype(str).replace(['nan', 'NAN', ''], 'sin crear')
    df_data['Estado'] = df_data['Estado'].astype(str).replace(['nan', 'NAN', ''], 'Programada')
    df_data['Sucursal'] = df_data['Sucursal'].astype(str).replace(['nan', 'NAN', 'sin crear'], '')
    df_data['Ciudad'] = df_data['Ciudad'].astype(str).replace(['nan', 'NAN'], '')
    
    # 6. INYECCIÓN DE METADATOS GLOBALES
    df_data['Cliente'] = cliente_actual
    df_data['NIT'] = nit_actual
    df_data['Ciudad_Global'] = ciudad_global_actual
    df_data['Fecha_Visita'] = fecha_actual
    df_data['Mantenimiento'] = df_data['Visita']
    
    # 7. ASIGNACIÓN DE KPIS (Semáforo Vectorizado con np.select)
    condiciones = [
        (df_data['OT'] == "sin crear"),
        (df_data['OT'].astype(str).str.startswith("OTT -", na=False)) & (df_data['Estado'] == "Programada"),
        (df_data['OT'].astype(str).str.startswith("OTT -", na=False)) & (df_data['Estado'].isin(["Cerrada", "Finalizada"]))
    ]
    df_data['Color_Semantico'] = np.select(condiciones, ["Rojo", "Amarillo", "Verde"], default="Desconocido")
    
    # Ordenar y retornar columnas requeridas por app.py
    columnas_finales = ['Cliente', 'NIT', 'Ciudad_Global', 'Fecha_Visita', 'Equipo', 'Mantenimiento', 'OT', 'Estado', 'Sucursal', 'Ciudad', 'Color_Semantico']
    return df_data[columnas_finales]
