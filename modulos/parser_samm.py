import pandas as pd
import numpy as np

def limpiar_reporte_samm(ruta_archivo):
    df_raw = pd.read_excel(ruta_archivo, header=None)
    
    datos_limpios = []
    
    # Memoria RAM del escáner (Contexto actual)
    cliente_actual = "DESCONOCIDO"
    fecha_actual = "DESCONOCIDA"
    nit_actual = "NO REGISTRA" # ¡NUEVO! Memoria para el NIT
    ciudad_global_actual = "NO REGISTRA" # ¡NUEVO! Memoria para la Ciudad del cliente
    
    col_idx = {"equipo": -1, "visita": -1, "ot": -1, "estado": -1, "sucursal": -1, "ciudad": -1}
    
    # Lista negra de basura
    palabras_ignoradas = ('EQUIPO', 'CLIENTE:', 'NIT:', 'CIUDAD:', 'DIRECCION:', 'CONTRATO', 'FECHA', 'NAN')
    
    # Escáner continuo (Lee el 100% del documento)
    for index, row in df_raw.iterrows():
        valores_meta = [str(x).strip() for x in row.values if pd.notna(x)]
        if not valores_meta:
            continue
            
        # 1. Detección de Metadatos
        if "CLIENTE:" in valores_meta:
            try: idx = valores_meta.index("CLIENTE:"); cliente_actual = valores_meta[idx + 1]
            except: pass
            
        # ¡NUEVO! Atrapamos el NIT y la CIUDAD global
        if "NIT:" in valores_meta:
            try: idx = valores_meta.index("NIT:"); nit_actual = valores_meta[idx + 1]
            except: pass
        if "CIUDAD:" in valores_meta:
            try: idx = valores_meta.index("CIUDAD:"); ciudad_global_actual = valores_meta[idx + 1]
            except: pass
                
        # 2. Detección de cambio de Fecha global
        for val in valores_meta:
            if val.startswith("Fecha Visita:"):
                fecha_actual = val.replace("Fecha Visita:", "").strip()
                
        # 3. Calibración del Radar
        if "Equipo" in valores_meta and ("Sucursal" in valores_meta or "Visita" in valores_meta):
            for i, val in enumerate(row.values):
                val_str = str(val).strip()
                if val_str == "Equipo": col_idx["equipo"] = i
                elif val_str == "Visita": col_idx["visita"] = i
                elif val_str == "OT": col_idx["ot"] = i
                elif val_str == "Estado": col_idx["estado"] = i
                elif val_str == "Sucursal": col_idx["sucursal"] = i
                elif val_str == "Ciudad": col_idx["ciudad"] = i 
            continue
            
        # 4. Cosecha de Datos de los Equipos
        if col_idx["equipo"] != -1:
            col0 = str(row.iloc[col_idx["equipo"]]).strip() if col_idx["equipo"] < len(row) else ""
            
            if col0 and not any(col0.upper().startswith(p) for p in palabras_ignoradas):
                equipo = col0
                visita = str(row.iloc[col_idx["visita"]]).strip() if col_idx["visita"] != -1 and pd.notna(row.iloc[col_idx["visita"]]) else ""
                ot = str(row.iloc[col_idx["ot"]]).strip() if col_idx["ot"] != -1 and pd.notna(row.iloc[col_idx["ot"]]) else "sin crear"
                estado = str(row.iloc[col_idx["estado"]]).strip() if col_idx["estado"] != -1 and pd.notna(row.iloc[col_idx["estado"]]) else "Programada"
                sucursal = str(row.iloc[col_idx["sucursal"]]).strip() if col_idx["sucursal"] != -1 and pd.notna(row.iloc[col_idx["sucursal"]]) else ""
                ciudad = str(row.iloc[col_idx["ciudad"]]).strip() if col_idx["ciudad"] != -1 and pd.notna(row.iloc[col_idx["ciudad"]]) else ""
                
                if ot.lower() == 'nan': ot = 'sin crear'
                if estado.lower() == 'nan': estado = 'Programada'
                if sucursal.lower() == 'nan' or sucursal.lower() == 'sin crear': sucursal = ''
                if ciudad.lower() == 'nan': ciudad = ''
                
                datos_limpios.append({
                    "Cliente": cliente_actual,
                    "NIT": nit_actual, # ¡NUEVO! Guardamos en tabla
                    "Ciudad_Global": ciudad_global_actual, # ¡NUEVO! Guardamos en tabla
                    "Fecha_Visita": fecha_actual,
                    "Equipo": equipo,
                    "Mantenimiento": visita,
                    "OT": ot,
                    "Estado": estado,
                    "Sucursal": sucursal,
                    "Ciudad": ciudad
                })

    df_final = pd.DataFrame(datos_limpios)
    
    # Asignación de KPIs
    if not df_final.empty:
        condiciones = [
            (df_final['OT'] == "sin crear"),
            (df_final['OT'].str.startswith("OTT -")) & (df_final['Estado'] == "Programada"),
            (df_final['OT'].str.startswith("OTT -")) & (df_final['Estado'].isin(["Cerrada", "Finalizada"]))
        ]
        opciones = ["Rojo", "Amarillo", "Verde"]
        df_final['Color_Semantico'] = np.select(condiciones, opciones, default="Desconocido")
    else:
        df_final['Color_Semantico'] = []
        
    return df_final