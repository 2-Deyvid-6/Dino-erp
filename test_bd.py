import pyodbc

# Credenciales directas
server = 'localhost'
database = 'sw_test'
username = 'bot_erp'
password = 'BotDinoPassword2026!'
driver = '{ODBC Driver 17 for SQL Server}'

# Forzamos un timeout de 5 segundos para que no se quede colgado
cadena_conexion = f'DRIVER={driver};SERVER={server},1433;DATABASE={database};UID={username};PWD={password};Encrypt=no;TrustServerCertificate=yes;'

print("⏳ Tocando la puerta del servidor SAMM...")

try:
    # timeout=5 aborta el intento si el servidor no responde en 5 segundos
    conn = pyodbc.connect(cadena_conexion, timeout=5)
    print("✅ ¡CONEXIÓN EXITOSA! El firewall y la red están perfectos.")
    
    # Hacemos una lectura de prueba ultra rápida
    cursor = conn.cursor()
    cursor.execute("SELECT TOP 3 equipo_codigo FROM view_equ_equipo")
    print("Datos de prueba traídos:")
    for row in cursor.fetchall():
        print(" -", row[0])
    
    conn.close()
    print("🎯 El problema es solo cómo Streamlit lee el archivo TOML.")

except Exception as e:
    print("❌ LA CONEXIÓN REBOTÓ. El error real es:")
    print(e)