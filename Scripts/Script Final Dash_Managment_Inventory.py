import win32com.client
import os
import time
from datetime import datetime

# ======= FOR GET STARTED OPTIONS ======= #

CARPETA_DESTINO = r"C:\Dashboard Downloads"
ARCHIVO_CSV = "LT24-L.csv"

RUTA_COMPLETA = os.path.join(CARPETA_DESTINO, ARCHIVO_CSV)

LAYOUT = "FELIXL-L"

# ======= LOG ======= #

def escribir_log(mensaje):

    archivo_log = os.path.join(
        CARPETA_DESTINO,
        "LT24_Automation_Log.txt"
    )

    with open(archivo_log, "a", encoding="utf-8") as log:

        fecha = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        log.write(f"{fecha} - {mensaje}\n")

# ======= FILE-1ER SHIFT ======= #

if not os.path.exists(CARPETA_DESTINO):
    os.makedirs(CARPETA_DESTINO)

escribir_log("Inicio de ejecución")

try:

    # ======= REPLACE CSV ======= #
 
    if os.path.exists(RUTA_COMPLETA):

        os.remove(RUTA_COMPLETA)

        escribir_log(
            "Archivo anterior eliminado"
        )

    # ======= CONECT TO SAP LUP ======= #

    SapGuiAuto = win32com.client.GetObject(
        "SAPGUI"
    )

    application = \
        SapGuiAuto.GetScriptingEngine

    connection = application.Children(0)

    session = connection.Children(0)

    escribir_log(
        "Conectado a SAP"
    )

    # ======= OPEN LT24 LUP SAP ======= #

    session.findById(
        "wnd[0]/tbar[0]/okcd"
    ).text = "/nLT24"

    session.findById(
        "wnd[0]"
    ).sendVKey(0)

    time.sleep(3)

    escribir_log(
        "LT24 abierto"
    )

    # ======= SAP LAYOPUT CORRECT ======= #
   
    try:

        session.findById(
            "wnd[0]/tbar[1]/btn[33]"
        ).press()

        time.sleep(2)

        session.findById(
            "wnd[1]/usr/txtGS_LAYOUT-VARIANT"
        ).text = LAYOUT

        session.findById(
            "wnd[1]/tbar[0]/btn[8]"
        ).press()

        escribir_log(
            f"Layout {LAYOUT} cargado"
        )

    except Exception:

        escribir_log(
            "No fue posible cargar Layout automáticamente. Cargando..."
        )

    time.sleep(2)

     # ======= F8 REPORT ======= #

    session.findById(
        "wnd[0]/tbar[1]/btn[8]"
    ).press()

    escribir_log(
        "Reporte ejecutado F8"
    )

    time.sleep(10)

     # ======= EXPORT CSV ======= #

    session.findById(
        "wnd[0]/mbar/menu[0]/menu[3]/menu[1]"
    ).select()

    time.sleep(3)

    session.findById(
        "wnd[1]/usr/ctxtDY_PATH"
    ).text = CARPETA_DESTINO

    session.findById(
        "wnd[1]/usr/ctxtDY_FILENAME"
    ).text = ARCHIVO_CSV

    session.findById(
        "wnd[1]/tbar[0]/btn[11]"
    ).press()

    escribir_log(
        "CSV exportado correctamente"
    )

    print(
        "Proceso completado correctamente"
    )

except Exception as error:

    escribir_log(
        f"ERROR: {str(error)}"
    )

    print(
        f"Error encontrado: {error}"
    )

finally:

    escribir_log(
        "Fin de ejecución"
    )
