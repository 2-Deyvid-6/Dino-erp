import pandas as pd
import io

# 1. Tu lista manual de equipos a combustión (Hardcodeada como solicitaste)
EQUIPOS_COMBUSTION = ['103', '105', '110', '123', '141', '168'] # Ejemplo, aquí pondrás todos

def procesar_horometros(df_samm, df_historico):
    # 1. Filtrar solo los equipos de combustión
    df_samm = df_samm[df_samm['Equipo'].isin(EQUIPOS_COMBUSTION)]
    
    # 2. Cruce de bases de datos (Merge)
    # Unimos la data actual de SAMM con tu archivo histórico usando el número de Equipo
    df_cruce = pd.merge(
        df_samm, 
        df_historico, 
        on='Equipo', 
        how='left'
    )
    
    # 3. Lógica de Estados (Semáforo de Filtros)
    def calcular_estado(row):
        # Aquí irá la lógica matemática. Ejemplo: si la diferencia es mayor a 250 horas
        diferencia = row['Horometro_Actual'] - row['Ultimo_Horometro']
        
        # NOTA: "repuestos pedidos" y "pendiente por instalar" requerirán una marca en SAMM o manual
        if diferencia >= 250:
            return "REQUIERE CAMBIO"
        else:
            return "FALTAN HORAS"
            
    df_cruce['Estado'] = df_cruce.apply(calcular_estado, axis=1)
    
    # 4. Seleccionar y ordenar las columnas exactas que pediste
    columnas_finales = [
        'Equipo', 'Modelo', 'Sucursal', 
        'Ultimo_Horometro', 'Fecha_Ultimo_Horometro', 
        'Horometro_Actual', 'Estado'
    ]
    df_final = df_cruce[columnas_finales]
    
    # 5. Segmentación de las 3 Tablas (Hojas de Excel)
    df_ajover = df_final[df_final['Cliente'].str.contains('AJOVER', na=False, case=False)]
    df_soberana_acesco = df_final[df_final['Cliente'].str.contains('SOBERANA|ACESCO', na=False, case=False)]
    df_otros = df_final[~df_final['Cliente'].str.contains('AJOVER|SOBERANA|ACESCO', na=False, case=False)]
    
    # 6. Generador del Archivo Excel en Memoria (Para descargar en Streamlit)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_ajover.to_excel(writer, sheet_name='AJOVER', index=False)
        df_soberana_acesco.to_excel(writer, sheet_name='SOBERANA Y ACESCO', index=False)
        df_otros.to_excel(writer, sheet_name='OTROS', index=False)
        
    return output.getvalue()