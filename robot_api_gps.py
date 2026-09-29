import requests
import pandas as pd
import time
import os
import json
from datetime import datetime, timedelta

# ==========================================
# ⚙️ CONFIGURACIÓN DEL ROBOT API
# ==========================================
USUARIO_GPS = "aux.sistemas@dinomontacargas.net"
PASSWORD_GPS = "Dino2026*" # Asegúrate de colocar la contraseña correcta
CARPETA_DESTINO = "datos_samm"

# Cargar flota desde config.json de forma dinámica
try:
    with open("config.json", "r", encoding="utf-8") as f:
        FLOTA_ACTIVA = json.load(f).get("flota_activa", [])
except:
    FLOTA_ACTIVA = []

def actualizar_historial_api(fecha_consulta=None):
    if not FLOTA_ACTIVA:
        print("⚠️ No hay vehículos en FLOTA_ACTIVA (Revisar config.json)")
        return False
        
    if not fecha_consulta:
        fecha_ref = datetime.now()
    else:
        fecha_ref = datetime.strptime(str(fecha_consulta), "%Y-%m-%d")
        
    fecha_str = fecha_ref.strftime('%Y-%m-%d')
    hora_inicio = f"{fecha_str} 00:00:00"

    # 1. Autenticación
    url_auth = "https://api.guardian.click/auth"
    payload = {"userName": USUARIO_GPS, "password": PASSWORD_GPS}
    
    resp_auth = requests.post(url_auth, data=payload)
    if resp_auth.status_code != 200:
        print(f"❌ Error Auth ({resp_auth.status_code})")
        return False
        
    token = resp_auth.json()['body']['token']
    headers_api = {'Authorization': f'Bearer {token}'}

    # 2. Obtener lista de equipos
    url_devices = "https://api.guardian.click/devices"
    resp_dev = requests.get(url_devices, headers=headers_api)
    if resp_dev.status_code != 200:
        if resp_dev.status_code == 429: time.sleep(60)
        return False
        
    datos_crudos = resp_dev.json()
    lista_equipos = datos_crudos.get('body', {}).get('devices', []) if isinstance(datos_crudos, dict) else datos_crudos
    if not isinstance(lista_equipos, list): lista_equipos = []
    
    mapa_ids = {str(eq.get('nombre_equipo', '')).strip().upper(): eq.get('id_equipo') for eq in lista_equipos if isinstance(eq, dict) and str(eq.get('nombre_equipo', '')).strip().upper() in FLOTA_ACTIVA}

    if not mapa_ids: return False

    registros_totales = []
    
    # 3. Descarga de Historial por Lotes Paginados
    for placa, id_equipo in mapa_ids.items():
        historial_completo = []
        fecha_peticion_inicio = hora_inicio
        
        for iteracion in range(30): 
            fecha_safe_inicio = str(fecha_peticion_inicio).replace(" ", "%20")
            
            url_hist = f"https://api.guardian.click/devices/history/{id_equipo}/{fecha_safe_inicio}/500?eventIds=0,71,81,255"
            resp_hist = requests.get(url_hist, headers=headers_api)
            
            if resp_hist.status_code == 200:
                hist_crudo = resp_hist.json()
                historial_parcial = []
                
                if isinstance(hist_crudo, list):
                    historial_parcial = hist_crudo
                elif isinstance(hist_crudo, dict):
                    body = hist_crudo.get('body', hist_crudo)
                    if isinstance(body, list):
                        historial_parcial = body
                    elif isinstance(body, dict):
                        if 'histories' in body:
                            for key_eq, datos_eq in body['histories'].items():
                                if isinstance(datos_eq, dict) and 'history' in datos_eq:
                                    historial_parcial.extend(datos_eq['history'])
                        elif 'history' in body:
                            historial_parcial = body['history']
                
                # Paginación inteligente
                if historial_parcial and isinstance(historial_parcial, list) and len(historial_parcial) > 0:
                    historial_completo.extend(historial_parcial)
                    
                    fechas_validas = [r.get('fecha') for r in historial_parcial if isinstance(r, dict) and r.get('fecha')]
                    if fechas_validas:
                        max_fecha_str = max(fechas_validas)[:19] 
                        try:
                            dt_last = datetime.strptime(max_fecha_str, "%Y-%m-%d %H:%M:%S")
                            nueva_fecha_inicio = (dt_last + timedelta(seconds=1)).strftime("%Y-%m-%d %H:%M:%S")
                            
                            if nueva_fecha_inicio == fecha_peticion_inicio:
                                break
                            fecha_peticion_inicio = nueva_fecha_inicio
                            time.sleep(0.3)
                            continue
                        except Exception:
                            break
                    else:
                        break
                else:
                    break 
            elif resp_hist.status_code == 429:
                print(f"   ⚠️ Límite de la API alcanzado. Esperando 10s...")
                time.sleep(10)
                continue
            else:
                print(f"   ❌ Error API {placa} ({resp_hist.status_code}): {resp_hist.text}")
                break
                
        if historial_completo:
            df_temp = pd.DataFrame(historial_completo)
            if 'fecha' in df_temp.columns:
                df_temp = df_temp.drop_duplicates(subset=['fecha']).sort_values('fecha')
                
            contador_util = 0 # 💡 ERROR SOLUCIONADO AQUÍ
            
            for _, row in df_temp.iterrows():
                evento_raw = str(row.get('descripcion_novedad', ''))
                tipo_nov = row.get('tipo_novedad')
                
                if tipo_nov not in [0, 71, 81, 255] and 'ignición' not in evento_raw.lower():
                    continue
                
                contador_util += 1
                
                # 💡 TRADUCCIÓN ELÉCTRICA CORREGIDA
                if tipo_nov == 81 or 'cerrada' in evento_raw.lower() or 'prendida' in evento_raw.lower():
                    evento_traducido = 'Ignición prendida'
                elif tipo_nov == 71 or 'abierta' in evento_raw.lower() or 'apagada' in evento_raw.lower():
                    evento_traducido = 'Ignición apagada'
                else:
                    evento_traducido = evento_raw

                registros_totales.append({
                    'Vehículo': placa,
                    'Fecha': row.get('fecha'),
                    'Evento': evento_traducido,
                    'Latitud': row.get('latitud'),
                    'Longitud': row.get('longitud'),
                    'Ubicación': '-'
                })
            
            print(f"   ✔️ {placa}: {contador_util} registros útiles extraídos.")

    # 4. Sobreescritura en disco (BLINDAJE ATÓMICO)
    if registros_totales:
        df_export = pd.DataFrame(registros_totales)
        if not os.path.exists(CARPETA_DESTINO): 
            os.makedirs(CARPETA_DESTINO)
            
        ruta_unica = os.path.join(CARPETA_DESTINO, "historial_gps.xlsx")
        ruta_temp = os.path.join(CARPETA_DESTINO, "temp_historial_gps.xlsx")
        
        # 1. Escribimos en el archivo temporal
        df_export.to_excel(ruta_temp, index=False)
        
        # 2. Reemplazamos atómicamente
        os.replace(ruta_temp, ruta_unica)
        
        print(f"✅ Excel MAESTRO sobreescrito de forma atómica y segura ({len(df_export)} filas útiles).")
        return True
    return False
