"""
Automatizacion SAP GUI - Descarga de anexos de OC y HES (con interfaz grafica)
================================================================================

Requisitos:
    pip install pywin32 openpyxl pywinauto
    (tkinter viene incluido con Python en Windows, no requiere instalacion aparte)

    - SAP GUI Scripting habilitado en cliente y servidor (ya confirmado
      funcionando con un usuario de prueba).
    - "Visualizar ventana de dialogo nativa de Microsoft Windows" activado
      en Opciones de SAP GUI > Accesibilidad y Scripting > Scripting
      (ya confirmado activado).
    - Debes tener SAP Logon abierto con una sesion YA CONECTADA antes de
      correr este script.

Hoja del Excel (agregado 30/08/2026):
    El archivo Excel puede variar de nombre y tener varias hojas, pero la
    hoja que siempre contiene los datos a procesar se llama, por defecto,
    "Plan de Descarga". Tanto la lectura como la escritura de resultados
    usan esa hoja por nombre (NOMBRE_HOJA), no la hoja activa del libro.
    El nombre de la hoja es configurable desde la interfaz grafica (campo
    "Nombre de hoja") por si en algun momento cambiara, pero viene
    precargado con "Plan de Descarga". Si el Excel no tiene una hoja con
    ese nombre, el proceso se detiene con un mensaje claro indicando las
    hojas disponibles en el archivo.

Estructura del Excel esperado (sin depender de nombres de encabezado):
    Por defecto: columna B -> bloque, columna D -> codigo de OC, columna
    H -> codigo(s) HES. Estos valores (columna Bloque, columna OC,
    columna HES, fila de inicio de datos) son CONFIGURABLES desde la
    interfaz grafica, asi que si el diseño del Excel varia entre archivos
    no hace falta tocar el codigo: se ajustan en los campos "Columna
    Bloque", "Columna OC", "Columna HES" y "Fila de inicio" de la ventana
    antes de iniciar la descarga. Las constantes COL_BLOQUE_LETTER,
    COL_OC_LETTER, COL_HES_LETTER y FILA_INICIO_DATOS de mas abajo solo
    se usan como valores por defecto que aparecen precargados en esos
    campos.

    IMPORTANTE (corregido 28/08/2026): si alguna de estas columnas tiene
    celdas combinadas (una sola celda de Bloque/OC/HES que visualmente
    abarca varias filas), openpyxl por defecto solo devuelve el valor en
    la celda ancla (la esquina superior izquierda del rango combinado) y
    None en el resto de celdas del rango. leer_filas_excel() resuelve el
    valor real de cualquier celda que pertenezca a un rango combinado
    antes de leerla, para las tres columnas (Bloque, OC y HES).

    IMPORTANTE (agregado 18/09/2026): una misma OC puede aparecer mas de
    una vez en filas NO contiguas del mismo bloque (por ejemplo separada
    por otras OC en el medio), cada vez con HES distintas. Ver el detalle
    en el docstring de leer_filas_excel() mas abajo.

Estructura de carpetas generada:
    <carpeta_destino>/<Bloque>/<OC>/            -> anexos de la propia OC
    <carpeta_destino>/<Bloque>/<OC>/<HES>/       -> anexos de cada HES de esa OC

ESCRITURA DE RESULTADOS EN EL MISMO EXCEL (agregado 30/08/2026):
    Ademas de descargar, el script escribe de vuelta en el propio Excel de
    entrada (en la hoja NOMBRE_HOJA) cuatro columnas por defecto: Y (Fecha
    de Descarga), Z (Responsable), AA (Observaciones) y AB (Comentario) -
    tambien configurables desde la interfaz grafica con los campos
    "Columna Fecha", "Columna Responsable", "Columna Observaciones" y
    "Columna Comentario", por si el diseño real del Excel varia.

    La Fecha y el Responsable se escriben tal cual el usuario los ingreso
    en la interfaz (campos "Fecha de descarga" y "Responsable"): el MISMO
    valor se aplica a todas las filas de TODAS las OC procesadas en esta
    corrida (no varian OC por OC).

    Las Observaciones si se calculan automaticamente por OC, en base a si
    la OC y cada una de sus HES terminaron con al menos un archivo
    (anexo) en su carpeta correspondiente:
      - Si la OC tiene archivos y TODAS sus HES tambien (o no tiene HES):
        "OK".
      - En cualquier otro caso, se arma una frase del tipo "OC contiene
        archivos, HES 111, 222 contienen archivos, pero HES 333 no tienen
        archivos" (ver _calcular_observacion() para el detalle exacto de
        cada combinacion).

    El Comentario (agregado 30/08/2026, reemplaza el guardado en .txt que
    tenia antes el script) se llena SOLO cuando, al recorrer la lista de
    anexos de una OC o de una HES, alguna fila resulta ser un
    comentario/nota de SAP en vez de un archivo exportable (ver detalle
    en el codigo, funcion _es_comentario()).

    Estas columnas se rellenan en las filas del Excel que pertenecen a
    esa OC (la fila donde aparece el codigo de OC y las filas de
    continuacion con sus HES), usando la misma logica de agrupacion que
    ya resuelve celdas combinadas y filas de continuacion.

    El Excel se guarda (sobrescribe el mismo archivo) despues de terminar
    cada OC, para no perder el progreso si el proceso se interrumpe a la
    mitad. Si el archivo esta abierto en Excel u otro programa al momento
    de guardar, el guardado fallara para esa OC (se registra un aviso en
    el log) pero el proceso continua igual con la siguiente OC; conviene
    tener el Excel CERRADO mientras corre el script (ahora esto se valida
    ANTES de arrancar, ver mejora "VALIDACION DE EXCEL ABIERTO" abajo).

PAUSA ENTRE OC (agregado 02/09/2026):
    Para que el ritmo de descarga no sea evidentemente mas rapido que el
    de una persona operando SAP manualmente, al terminar de procesar cada
    OC (despues de guardar el Excel con sus resultados) el proceso espera
    PAUSA_ENTRE_OC segundos antes de pasar a la siguiente OC. Este valor
    es configurable desde la interfaz grafica (campo "Pausa entre OC
    (segundos)"), precargado en 60 segundos por defecto. No se aplica
    despues de la ULTIMA OC del Excel, ni cuando una OC se omite por
    reanudacion. La espera se hace en tramos cortos (de a 1 segundo) en
    vez de un solo time.sleep() largo, para que el log vaya mostrando una
    cuenta regresiva y para poder interrumpirla si el usuario aprieta
    "Detener".

================================================================================
MEJORAS AGREGADAS (07/09/2026), sobre la version que ya funcionaba en
produccion. Cada una se explica en detalle junto al codigo que la
implementa, aqui va solo un resumen:

  1) REANUDACION: antes de procesar cada OC (y cada HES dentro de una OC
     que no este completa), el script revisa si ya hay archivos
     descargados en la carpeta correspondiente. Si una OC completa (ella
     y todas sus HES) ya tiene archivos, se omite por completo -sin tocar
     SAP para nada- y solo se recalculan Fecha/Responsable/Observacion en
     el Excel. Esto permite cortar el proceso a mitad de camino (por un
     corte de luz, un error, o simplemente porque se quiere pausar) y
     volver a correrlo desde el principio sin que vuelva a descargar lo
     que ya estaba bien. Se puede desactivar destildando la casilla
     "Reanudar" en la interfaz, para forzar una descarga completa desde
     cero.

  2) BOTON DETENER: ahora hay un boton "Detener" que se habilita mientras
     el proceso corre. Al presionarlo, se activa un threading.Event que
     el proceso revisa entre OC, entre HES y durante la pausa entre OC;
     al detectarlo, termina de forma ordenada (sin dejar SAP a mitad de
     una operacion) en vez de tener que forzar el cierre de la ventana.

  3) LOG A ARCHIVO: ademas de mostrarse en pantalla, cada corrida escribe
     su log completo (con hora de cada linea) a un archivo de texto
     dentro de la carpeta de descargas, en una subcarpeta "logs". Sirve
     de auditoria de que paso en cada corrida, incluso si se cerro la
     ventana antes de poder leer todo en pantalla.

  4) CONFIGURACION PERSISTENTE: los valores de la interfaz (carpeta de
     descargas, nombre de hoja, columnas, responsable, pausa, casilla de
     reanudacion) se guardan en un archivo de configuracion en la carpeta
     personal del usuario y se recargan automaticamente la proxima vez
     que se abra el programa, para no tener que volver a tipear todo cada
     vez.

  5) BACKUP DEL EXCEL: antes de empezar a escribir resultados, se guarda
     una copia del Excel original (con marca de tiempo) en la misma
     carpeta donde esta el archivo, como red de seguridad por si algo
     saliera mal durante el guardado.

  6) VERIFICACION DE SESION SAP: antes de procesar cada OC se comprueba
     que la sesion de SAP siga respondiendo. Si se perdio la conexion
     (por ejemplo, si SAP se cerro o se desconecto), el script intenta
     reconectar automaticamente unas cuantas veces; si no lo logra,
     detiene el proceso con un mensaje claro en vez de seguir fallando
     OC tras OC con errores confusos.

  7) NOTIFICACION AL TERMINAR: al finalizar el proceso (ya sea completo,
     detenido por el usuario, o cortado por un error), se muestra un
     aviso emergente y un sonido, para no tener que estar mirando la
     ventana mientras corre.

  8) PROGRESO Y VALIDACIONES: la interfaz ahora muestra "OC X de Y" en
     tiempo real, y se valida ANTES de arrancar que el Excel no este
     abierto en otro programa (evita perder tiempo y descubrir el
     problema recien despues de la primera OC).

Todos los IDs de SAP y la logica de scripting (deadlock de
pressToolbarButton, backend win32 para el dialogo de exportar, deteccion
de comentarios, espera adaptativa para archivos pesados, reintentos ante
"SAP ocupado", recuperacion de sesion) se mantienen EXACTAMENTE igual que
en la version que ya estaba confirmada funcionando - estas mejoras son
todas aditivas, no tocan esa logica.
================================================================================

NOTA IMPORTANTE SOBRE EL IMPORT DE pywinauto:
    "from pywinauto import Desktop" se hace DENTRO de la funcion que lo
    necesita (descargar_todos_los_anexos), no al inicio del archivo.
    Si se importa al inicio, pywinauto inicializa COM en un modo que
    choca con el dialogo nativo de Tkinter (filedialog.askdirectory /
    askopenfilename), y el selector de carpeta/archivo se queda
    congelado con el icono de carga infinito. Al importarlo solo cuando
    se usa (ya con los selectores de la interfaz cerrados), se evita
    el conflicto.

NOTA SOBRE EL BACKEND DE pywinauto PARA EL DIALOGO 'Exportar fichero':
    Confirmado por diagnostico (28/08/2026, SAP GUI 760):
    el dialogo 'Exportar fichero' es clase #32770 (dialogo comun clasico
    de Windows). backend="uia" NO lo detecta en este entorno, pero
    backend="win32" si. Se descarto un problema de privilegios/UIPI
    (saplogon.exe corre sin elevacion, igual que el script). Por eso
    todo el manejo de este dialogo usa exclusivamente backend="win32".

NOTA CRITICA SOBRE EL DEADLOCK DE pressToolbarButton("%ATTA_EXPORT"):
    Confirmado por comportamiento observado (28/08/2026): la llamada COM
    session.findById(...).pressToolbarButton("%ATTA_EXPORT") NO RETORNA
    hasta que el dialogo nativo "Exportar fichero" que ella misma abre
    se cierra. Esto es porque SAP GUI entra en su propio bucle modal de
    Windows internamente y no puede responder a la llamada RPC hasta que
    ese bucle termina. Si se llama esta funcion normalmente (en el mismo
    hilo que luego intenta manejar el dialogo), se produce un DEADLOCK.
    SOLUCION: pressToolbarButton("%ATTA_EXPORT") se dispara en un hilo
    aparte (_disparar_exportar), dejando el hilo principal libre para
    buscar y manejar el dialogo con pywinauto mientras esa llamada COM
    sigue bloqueada esperando. Una vez el dialogo se cierra (al hacer
    click en "Guardar"), la llamada en el hilo aparte retorna y se hace
    join() sobre ese hilo antes de continuar con la siguiente fila.

NOTA SOBRE DETECCION DE COMENTARIOS (actualizado 30/08/2026):
    Cuando exportar una fila de la grilla de anexos falla, puede ser
    porque esa fila no es un archivo sino un comentario/nota de SAP (que
    no tiene "Exportar fichero"). _es_comentario() lo verifica haciendo
    doble clic sobre la fila: si eso abre una ventana emergente (wnd[2]),
    se confirma que era un comentario, se cierra esa ventana y se avisa
    hacia arriba. Solo interesa saber que existio un comentario, para
    anotarlo en la columna Comentario del Excel (no se guarda su texto).

NOTA SOBRE ARCHIVOS PESADOS QUE TARDAN EN EXPORTARSE (agregado 31/08/2026):
    _esperar_exportacion_terminada() espera de forma adaptativa (usando
    session.Busy y verificacion en disco de que el archivo dejo de
    crecer) en vez de un sleep fijo, para no chocar con el error COM
    -2147483638 ("El dato necesario para completar esta operación no
    está disponible todavía") al pasar a la fila siguiente demasiado
    pronto.

NOTA SOBRE REINTENTOS ANTE "OPERACION NO DISPONIBLE TODAVIA" (agregado 31/08/2026):
    Toda llamada de scripting que pueda chocar con SAP todavia ocupado
    pasa por _llamar_con_reintentos(): si falla con el error COM
    -2147483638, se espera y se reintenta la MISMA llamada, hasta
    REINTENTOS_SAP_OCUPADO veces, en vez de propagar el error de
    inmediato (lo que antes dejaba ventanas colgadas y hacia fallar en
    cascada las filas siguientes).

NOTA SOBRE RECUPERACION DE SESION (agregado 31/08/2026):
    _recuperar_sesion() se llama al terminar de procesar CADA OC y CADA
    HES: cierra cualquier ventana que haya quedado abierta colgada, para
    que un problema puntual en una fila no se arrastre a las filas
    siguientes del Excel.
"""

import os
import re
import json
import time
import shutil
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, scrolledtext, messagebox

import openpyxl
import win32com.client

ESPERA = 1.0  # segundos entre acciones, sube si tu SAP responde lento

# Tiempos de espera tras confirmar la exportacion de un anexo (ver nota
# "ARCHIVOS PESADOS QUE TARDAN EN EXPORTARSE" arriba).
TIMEOUT_ESPERA_BUSY = 120     # segundos maximos esperando a que SAP deje de estar "ocupado"
TIMEOUT_ESPERA_ARCHIVO = 60   # segundos maximos esperando a que el archivo aparezca/se estabilice en disco

# Reintentos ante el error transitorio "El dato necesario para completar
# esta operación no está disponible todavía" (ver nota "REINTENTOS ANTE
# OPERACION NO DISPONIBLE TODAVIA" mas arriba).
REINTENTOS_SAP_OCUPADO = 20
ESPERA_ENTRE_REINTENTOS = 3.0  # segundos

# Pausa entre OC, para simular ritmo manual (ver nota "PAUSA ENTRE OC" arriba).
PAUSA_ENTRE_OC = 60  # segundos

NOMBRE_HOJA = "Plan de Descarga"

COL_BLOQUE_LETTER = "B"
COL_OC_LETTER = "D"
COL_HES_LETTER = "H"
FILA_INICIO_DATOS = 2  # fila 1 = encabezados

# Columnas donde se escriben los resultados de vuelta en el mismo Excel
COL_FECHA_LETTER = "Y"
COL_RESPONSABLE_LETTER = "Z"
COL_OBSERVACIONES_LETTER = "AA"
COL_COMENTARIO_LETTER = "AB"

# Reintentos al intentar reconectar la sesion SAP si se detecta que se
# perdio la conexion durante una corrida larga (mejora 6, ver mas abajo).
REINTENTOS_RECONEXION_SAP = 3
ESPERA_ENTRE_RECONEXIONES = 5.0  # segundos

# Archivo de configuracion persistente (mejora 4, ver mas abajo). Se
# guarda en la carpeta personal del usuario para que sobreviva a
# reinstalaciones del script y no dependa de la carpeta desde la que se
# corre.
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".sap_descarga_oc_hes_config.json")

# Campos de la interfaz que se guardan y recargan automaticamente entre
# sesiones. Deliberadamente NO se incluyen la ruta del Excel (cambia
# seguido de corrida en corrida) ni la Fecha de descarga (es especifica
# del dia), para que no queden valores viejos precargados por error.
CAMPOS_CONFIG_PERSISTENTE = [
    "carpeta_var",
    "hoja_var",
    "col_bloque_var",
    "col_oc_var",
    "col_hes_var",
    "fila_inicio_var",
    "col_fecha_var",
    "col_responsable_var",
    "col_observaciones_var",
    "col_comentario_var",
    "responsable_var",
    "pausa_var",
    "reanudar_var",
]


# ---------------------------------------------------------------------------
# CONEXION A LA SESION SAP YA ABIERTA
# ---------------------------------------------------------------------------
def get_sap_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def _es_error_sap_ocupado(e):
    """True si la excepcion corresponde al error transitorio de SAP
    ('El dato necesario para completar esta operación no está
    disponible todavía', hresult -2147483638), que aparece cuando se le
    pide a SAP que haga algo mientras todavia esta terminando de
    procesar la exportacion de un archivo pesado en segundo plano.
    Ver nota "REINTENTOS ANTE OPERACION NO DISPONIBLE TODAVIA" al inicio
    del archivo."""
    try:
        if e.args and e.args[0] == -2147483638:
            return True
    except Exception:
        pass
    texto = str(e)
    return "no está disponible todavía" in texto or "no esta disponible todavia" in texto


def _llamar_con_reintentos(func, log, contexto, intentos=REINTENTOS_SAP_OCUPADO, espera=ESPERA_ENTRE_REINTENTOS):
    """Ejecuta func() reintentando automaticamente si falla con el error
    transitorio de SAP ocupado (ver _es_error_sap_ocupado). Cualquier
    otro tipo de excepcion se relanza de inmediato, sin reintentar."""
    ultimo_error = None
    for intento in range(1, intentos + 1):
        try:
            return func()
        except Exception as e:
            ultimo_error = e
            if not _es_error_sap_ocupado(e):
                raise
            log(
                f"      ({contexto}) SAP todavia esta terminando de procesar "
                f"(intento {intento}/{intentos}), reintentando en {espera}s..."
            )
            time.sleep(espera)
    raise ultimo_error


def _recuperar_sesion(session, log, max_wnd=6):
    """Red de seguridad para dejar la sesion en un estado conocido:
    cierra cualquier ventana secundaria (wnd[1], wnd[2], ...) que haya
    quedado abierta colgada, de la de mayor indice a la menor."""
    for idx in range(max_wnd, 0, -1):
        try:
            wnd = session.findById(f"wnd[{idx}]")
        except Exception:
            continue
        try:
            wnd.close()
            log(f"      (recuperacion: se cerro una ventana wnd[{idx}] que habia quedado abierta)")
            time.sleep(0.5)
        except Exception:
            pass


def _pausar_entre_oc(segundos, log, detener_event=None):
    """Espera 'segundos' antes de pasar a la siguiente OC, para que el
    ritmo de descarga no se note evidentemente mas rapido que el de una
    persona operando SAP manualmente. Se hace en tramos de 1 segundo (en
    vez de un unico time.sleep largo) para poder ir mostrando una cuenta
    regresiva en el log, y para poder interrumpirla de inmediato si el
    usuario aprieta "Detener" (mejora 2) en vez de tener que esperar a
    que termine la pausa completa."""
    if segundos <= 0:
        return
    log(f"    (pausa de {segundos}s antes de la siguiente OC...)")
    restante = segundos
    while restante > 0:
        if detener_event is not None and detener_event.is_set():
            log("      (pausa interrumpida: se solicito detener el proceso)")
            return
        paso = min(10, restante)
        time.sleep(paso)
        restante -= paso
        if restante > 0:
            log(f"      ...faltan {restante}s")
    log("      (pausa terminada, continuando)")


# ---------------------------------------------------------------------------
# MEJORA 6: VERIFICACION Y RECONEXION DE LA SESION SAP
# ---------------------------------------------------------------------------
def _sesion_esta_viva(session):
    """Comprueba de forma barata si la sesion SAP todavia responde,
    consultando una propiedad simple. Si SAP se cerro, se desconecto, o
    la sesion quedo en un estado invalido, esto lanza una excepcion COM
    y devolvemos False en vez de dejar que el error se propague sin
    contexto mas adelante, en medio del procesamiento de una OC."""
    try:
        _ = session.Children.Count
        return True
    except Exception:
        return False


def _reconectar_sesion(log, intentos=REINTENTOS_RECONEXION_SAP, espera=ESPERA_ENTRE_RECONEXIONES):
    """Intenta reconectar con una sesion SAP activa. Se usa cuando
    _sesion_esta_viva() detecta que se perdio la conexion durante una
    corrida larga (por ejemplo, si SAP se cerro solo, o si la maquina se
    quedo sin recursos). Devuelve la nueva sesion si lo logra, o None si
    se agotan los intentos - en ese caso el llamador debe detener el
    proceso en vez de seguir intentando OC por OC con una sesion muerta."""
    for intento in range(1, intentos + 1):
        try:
            nueva_sesion = get_sap_session()
            if _sesion_esta_viva(nueva_sesion):
                log(f"    (conexion con SAP recuperada, intento {intento}/{intentos})")
                return nueva_sesion
        except Exception as e:
            log(f"    (fallo al reconectar con SAP, intento {intento}/{intentos}: {e})")
        time.sleep(espera)
    return None


# ---------------------------------------------------------------------------
# NAVEGACION
# ---------------------------------------------------------------------------
def volver_pantalla_inicial(session, log):
    """Envia F3 ('Atras') sobre la ventana principal.

    CONFIRMADO por grabacion completa (Scriptcompleto.vbs, 28/08/2026):
    despues de descargar los anexos de un documento (OC o HES) y cerrar
    la ventana emergente de anexos, la grabacion manda sendVKey 3 sobre
    wnd[0] ANTES de escribir la siguiente transaccion en okcd. Sin este
    paso, la pantalla de visualizacion del documento anterior queda a
    medias y el siguiente cambio de transaccion (via okcd) deja el
    estado inconsistente, lo que produce fallos de tipo
    "The control could not be found by id." en el documento siguiente.

    Usa _llamar_con_reintentos() porque si SAP todavia esta terminando de
    procesar una exportacion pesada, el sendVKey puede fallar con el
    error transitorio de "SAP ocupado" - en ese caso se reintenta en vez
    de propagar el error de inmediato.
    """
    _llamar_con_reintentos(
        lambda: session.findById("wnd[0]").sendVKey(3), log, "volver a pantalla inicial"
    )


def abrir_oc(session, oc, log):
    """Abre la OC en ME23N (secuencia confirmada por Scriptcompleto.vbs)."""
    log(f"    [OC {oc}] maximizando ventana y navegando a ME23N...")
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ME23N"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(ESPERA)

    log(f"    [OC {oc}] presionando boton 'Otro documento' (btn[17])...")
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    time.sleep(ESPERA)

    log(f"    [OC {oc}] escribiendo codigo de OC en el dialogo...")
    campo_oc = session.findById(
        "wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN"
    )
    campo_oc.text = oc
    campo_oc.caretPosition = len(oc)
    session.findById("wnd[1]").sendVKey(0)
    time.sleep(ESPERA)


def abrir_hes(session, hes, log):
    """Abre la HES en ML81N (secuencia confirmada por Scriptcompleto.vbs, 28/08/2026)."""
    log(f"    [HES {hes}] maximizando ventana y navegando a ML81N...")
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ML81N"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(ESPERA)

    log(f"    [HES {hes}] presionando boton 'Otro documento' (btn[17])...")
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    time.sleep(ESPERA)

    log(f"    [HES {hes}] escribiendo codigo de HES en el dialogo...")
    campo_hes = session.findById("wnd[1]/usr/ctxtRM11R-LBLNI")
    campo_hes.text = hes
    campo_hes.setFocus()
    campo_hes.caretPosition = len(hes)
    session.findById("wnd[1]").sendVKey(0)
    time.sleep(ESPERA)


def _ventana_anexos_existe(session, timeout=3, intervalo=0.2):
    """Sondea un rato corto si la ventana emergente de anexos (con la
    grilla de anexos adentro) realmente llego a abrirse."""
    fin = time.time() + timeout
    while time.time() < fin:
        try:
            session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell")
            return True
        except Exception:
            time.sleep(intervalo)
    return False


def abrir_lista_anexos(session, log):
    """Abre la lista de anexos via 'Servicios para objeto'.

    Devuelve False (sin lanzar excepcion) tanto si el documento no tiene
    anexos como si la opcion no esta disponible; devuelve True solo si la
    ventana de anexos se abrio de verdad.
    """
    try:
        session.findById("wnd[0]/titl/shellcont/shell").pressContextButton("%GOS_TOOLBOX")
        session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem("%GOS_VIEW_ATTA")
    except Exception as e:
        log(f"      (no se pudo abrir anexos - probablemente el documento no tiene: {e})")
        return False

    if not _ventana_anexos_existe(session):
        log("      (opcion de anexos no disponible - documento sin anexos)")
        return False

    time.sleep(ESPERA)
    return True


def _nombre_sin_colision(carpeta_destino, nombre_archivo):
    """Si nombre_archivo ya existe en carpeta_destino, le agrega _1, _2, etc.
    antes de la extension para no sobreescribir un anexo distinto que
    tenga casualmente el mismo nombre."""
    # El nombre proviene de SAP: quitar separadores y caracteres especiales
    # para que un anexo no pueda salir de la carpeta elegida.
    nombre_archivo = os.path.basename(nombre_archivo.replace('\\', '/'))
    nombre_archivo = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', nombre_archivo).strip(' .')
    if not nombre_archivo:
        nombre_archivo = 'anexo'
    ruta = os.path.normpath(os.path.join(carpeta_destino, nombre_archivo))
    if not os.path.exists(ruta):
        return ruta

    base, ext = os.path.splitext(nombre_archivo)
    contador = 1
    while True:
        candidato = os.path.normpath(
            os.path.join(carpeta_destino, f"{base}_{contador}{ext}")
        )
        if not os.path.exists(candidato):
            return candidato
        contador += 1


# ---------------------------------------------------------------------------
# MANEJO DEL DIALOGO NATIVO "Exportar fichero" (backend win32, ver nota arriba)
# ---------------------------------------------------------------------------
def _esperar_dialogo_exportar(timeout=20, intervalo=0.3):
    from pywinauto import Desktop

    fin = time.time() + timeout
    ultimo_error = None
    while time.time() < fin:
        try:
            dlg = Desktop(backend="win32").window(
                title="Exportar fichero", class_name="#32770"
            )
            if dlg.exists() and dlg.is_visible():
                dlg.set_focus()
                return dlg
        except Exception as e:
            ultimo_error = e
        time.sleep(intervalo)

    raise TimeoutError(
        f"No se encontro el dialogo 'Exportar fichero' tras {timeout}s. "
        f"Ultimo error: {ultimo_error}"
    )


def _ubicar_campo_nombre(dlg):
    try:
        combo = dlg.child_window(class_name="ComboBoxEx32", found_index=0)
        return combo.child_window(class_name="Edit")
    except Exception:
        pass

    edits = dlg.children(class_name="Edit")
    if not edits:
        raise RuntimeError("No se encontro ningun campo Edit en el dialogo.")
    return edits[-1]


def _confirmar_guardar(dlg, campo_nombre, log, i):
    try:
        campo_nombre.set_focus()
    except Exception:
        pass
    time.sleep(0.2)
    campo_nombre.type_keys("{ENTER}", pause=0.05)
    time.sleep(0.5)

    if not dlg.exists():
        return

    log(f"      [fila {i}] ENTER no cerro el dialogo, buscando boton Guardar...")
    boton_guardar = dlg.child_window(
        class_name="Button", title_re=".*[Gg]uardar.*|.*[Ss]ave.*"
    )
    boton_guardar.click_input()


def _esperar_exportacion_terminada(session, ruta_completa, log, i):
    """Espera a que SAP realmente termine de exportar el anexo antes de
    seguir con la siguiente fila (ver nota "ARCHIVOS PESADOS QUE TARDAN
    EN EXPORTARSE" al inicio del archivo)."""
    fin_busy = time.time() + TIMEOUT_ESPERA_BUSY
    try:
        while True:
            if not session.Busy:
                break
            if time.time() >= fin_busy:
                log(
                    f"      [fila {i}] aviso: SAP siguio 'ocupado' mas de "
                    f"{TIMEOUT_ESPERA_BUSY}s exportando, se continua igual."
                )
                break
            time.sleep(0.5)
    except Exception as e:
        log(
            f"      [fila {i}] aviso: no se pudo consultar session.Busy ({e}), "
            f"se sigue solo con la verificacion en disco."
        )

    fin_archivo = time.time() + TIMEOUT_ESPERA_ARCHIVO
    tam_anterior = -1
    estable = False
    while time.time() < fin_archivo:
        if os.path.isfile(ruta_completa):
            tam_actual = os.path.getsize(ruta_completa)
            if tam_actual == tam_anterior and tam_actual > 0:
                estable = True
                break
            tam_anterior = tam_actual
        time.sleep(1.0)

    if not os.path.isfile(ruta_completa):
        log(
            f"      [fila {i}] aviso: tras esperar, no se encontro el archivo "
            f"en '{ruta_completa}'."
        )
    elif not estable:
        log(
            f"      [fila {i}] aviso: el archivo aparecio en '{ruta_completa}' "
            f"pero su tamaño seguia cambiando al agotarse la espera."
        )


def _cerrar_dialogo_si_quedo_abierto(log):
    from pywinauto import Desktop
    try:
        dlg = Desktop(backend="win32").window(
            title="Exportar fichero", class_name="#32770"
        )
        if dlg.exists():
            dlg.close()
            log("      (dialogo colgado cerrado tras el error)")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# COMENTARIOS EN LA LISTA DE ANEXOS (no son archivos, no se pueden exportar)
# ---------------------------------------------------------------------------
def _es_comentario(session, grid, fila_idx, log):
    """Verifica si la fila fila_idx de la grilla de anexos es un
    comentario/nota (no un archivo exportable)."""
    try:
        grid.currentCellColumn = "BITM_DESCR"
        grid.selectedRows = str(fila_idx)
        grid.doubleClickCurrentCell()
    except Exception as e:
        log(f"      [fila {fila_idx}] no se pudo hacer doble clic para revisar si es comentario: {e}")
        return False

    time.sleep(ESPERA)

    try:
        ventana_comentario = session.findById("wnd[2]")
    except Exception:
        log(f"      [fila {fila_idx}] no se abrio ninguna ventana al hacer doble clic (no era un comentario reconocible).")
        return False

    try:
        ventana_comentario.close()
    except Exception:
        pass

    time.sleep(0.3)
    log(f"      [fila {fila_idx}] era un comentario (no un archivo). Se anotara en la columna Comentario del Excel.")
    return True


def descargar_todos_los_anexos(session, carpeta_destino, log):
    """Descarga todos los anexos de la ventana emergente actual a carpeta_destino.

    Devuelve True si en esta lista de anexos se detecto al menos un
    comentario/nota (ver _es_comentario), False en caso contrario.
    """
    from pywinauto import Desktop  # noqa: F401 (usado en los helpers de arriba)

    os.makedirs(carpeta_destino, exist_ok=True)

    grid = session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell")
    total_filas = grid.RowCount

    if total_filas == 0:
        log("      (sin anexos)")
        session.findById("wnd[1]").close()
        return False

    hubo_comentario = False

    for i in range(total_filas):
        grid.setCurrentCell(i, "BITM_DESCR")
        grid.selectedRows = str(i)

        error_boton = {}

        def _disparar_exportar():
            try:
                grid.pressToolbarButton("%ATTA_EXPORT")
            except Exception as e:
                error_boton["error"] = e

        hilo_boton = threading.Thread(target=_disparar_exportar, daemon=True)
        hilo_boton.start()

        try:
            log(f"      [fila {i}] esperando dialogo 'Exportar fichero'...")
            dlg = _esperar_dialogo_exportar(timeout=10)

            log(f"      [fila {i}] dialogo encontrado, ubicando campo de nombre...")
            campo_nombre = _ubicar_campo_nombre(dlg)

            nombre_original = campo_nombre.window_text().strip()
            log(f"      [fila {i}] nombre detectado: {nombre_original or '(vacio)'}")

            if nombre_original:
                ruta_completa = _nombre_sin_colision(carpeta_destino, nombre_original)
            else:
                ruta_completa = os.path.normpath(
                    os.path.join(carpeta_destino, f"anexo_{i + 1}")
                )

            campo_nombre.set_edit_text(ruta_completa)
            time.sleep(0.3)

            log(f"      [fila {i}] guardando en: {ruta_completa}")
            _confirmar_guardar(dlg, campo_nombre, log, i)

            hilo_boton.join(timeout=10)
            if error_boton.get("error"):
                raise error_boton["error"]

            log(f"      [fila {i}] esperando a que SAP termine de exportar el archivo...")
            _esperar_exportacion_terminada(session, ruta_completa, log, i)
            log(f"      [fila {i}] OK")
        except Exception as e:
            _cerrar_dialogo_si_quedo_abierto(log)
            hilo_boton.join(timeout=5)

            if _es_comentario(session, grid, i, log):
                hubo_comentario = True
            else:
                log(f"      Aviso: no se pudo completar la exportacion en fila {i}: {e}")

    try:
        _llamar_con_reintentos(
            lambda: session.findById("wnd[1]").close(), log, "cerrar lista de anexos"
        )
    except Exception as e:
        log(
            f"      Aviso: no se pudo cerrar la lista de anexos tras varios "
            f"reintentos ({e}); se intentara recuperar la sesion mas adelante."
        )
    return hubo_comentario


# ---------------------------------------------------------------------------
# RESULTADO POR OC (para escribir Fecha / Responsable / Observaciones)
# ---------------------------------------------------------------------------
def _carpeta_tiene_archivos(carpeta):
    """True si 'carpeta' existe y contiene al menos un archivo directamente
    dentro de ella (no cuenta subcarpetas ni su contenido)."""
    if not carpeta or not os.path.isdir(carpeta):
        return False
    for nombre in os.listdir(carpeta):
        if os.path.isfile(os.path.join(carpeta, nombre)):
            return True
    return False


def _calcular_observacion(oc_tiene_archivos, hes_list, hes_tiene_archivos):
    """Arma el texto de Observaciones para una OC, segun si ella y cada
    una de sus HES terminaron con archivos descargados."""
    hes_ok = [h for h in hes_list if hes_tiene_archivos.get(h)]
    hes_bad = [h for h in hes_list if not hes_tiene_archivos.get(h)]

    if oc_tiene_archivos and not hes_bad:
        return "OK"

    oc_texto = "OC contiene archivos" if oc_tiene_archivos else "OC no contiene archivos"

    if hes_ok and hes_bad:
        return (
            f"{oc_texto}, HES {', '.join(hes_ok)} contienen archivos, "
            f"pero HES {', '.join(hes_bad)} no tienen archivos"
        )
    if hes_ok:
        return f"{oc_texto}, HES {', '.join(hes_ok)} contienen archivos"
    if hes_bad:
        return f"{oc_texto}, HES {', '.join(hes_bad)} no tienen archivos"
    return oc_texto


# ---------------------------------------------------------------------------
# MEJORA 1: REANUDACION - deteccion de OC / HES ya descargadas
# ---------------------------------------------------------------------------
def _oc_esta_completa_en_disco(carpeta_oc, hes_list):
    """True si la OC ya tiene archivos descargados en su carpeta y,
    ademas, TODAS sus HES (si tiene alguna) tambien ya tienen archivos en
    la suya. Se usa para decidir si se puede omitir por completo el
    procesamiento de esta OC en SAP (mejora 1: reanudacion)."""
    if not _carpeta_tiene_archivos(carpeta_oc):
        return False
    for hes in hes_list:
        carpeta_hes = os.path.join(carpeta_oc, hes)
        if not _carpeta_tiene_archivos(carpeta_hes):
            return False
    return True


# ---------------------------------------------------------------------------
# LECTURA DEL EXCEL (por columna, no por nombre de encabezado)
# ---------------------------------------------------------------------------
def leer_filas_excel(
    ruta_excel,
    nombre_hoja=NOMBRE_HOJA,
    col_bloque_letter=COL_BLOQUE_LETTER,
    col_oc_letter=COL_OC_LETTER,
    col_hes_letter=COL_HES_LETTER,
    fila_inicio=FILA_INICIO_DATOS,
    log=None,
):
    """Lee los Bloques, sus OC y las HES de cada OC desde el Excel, en la
    hoja nombre_hoja (por defecto "Plan de Descarga").

    Devuelve una lista de tuplas
        (bloque, [(oc, hes_list, filas, fila_principal, mapa_hes_fila), ...])
    donde:
      - 'filas' es la lista de numeros de fila del Excel (1-indexados)
        que pertenecen a esa OC.
      - 'fila_principal' es la primera fila de ese grupo (donde aparece
        el codigo de la OC POR PRIMERA VEZ en el bloque).
      - 'mapa_hes_fila' es un dict {codigo_hes: fila} con la fila donde
        aparece cada HES de esa OC.

    NOTA SOBRE OC REPETIDAS DENTRO DEL MISMO BLOQUE (agregado 18/09/2026):
    una OC puede aparecer mas de una vez en filas NO contiguas del mismo
    bloque (por ejemplo, separada por otras OC en el medio), cada vez con
    HES distintas. En ese caso todas las apariciones se agrupan en el
    MISMO grupo (misma carpeta de OC, mismo 'fila_principal' = la primera
    aparicion): no se vuelve a procesar la OC en SAP una segunda vez,
    solo se le suman las HES nuevas. Si la misma OC aparece en bloques
    DIFERENTES, cada bloque la trata como una OC independiente (carpetas
    Bloque/OC distintas), sin ningun cambio de comportamiento ahi.
    """
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    if nombre_hoja not in wb.sheetnames:
        raise ValueError(
            f"El Excel no tiene una hoja llamada '{nombre_hoja}'. "
            f"Hojas disponibles: {', '.join(wb.sheetnames)}"
        )
    ws = wb[nombre_hoja]

    col_bloque = openpyxl.utils.column_index_from_string(col_bloque_letter.strip().upper())
    col_oc = openpyxl.utils.column_index_from_string(col_oc_letter.strip().upper())
    col_hes = openpyxl.utils.column_index_from_string(col_hes_letter.strip().upper())

    valores_merge = {}
    for rango in ws.merged_cells.ranges:
        valor_ancla = ws.cell(row=rango.min_row, column=rango.min_col).value
        for fila_r in range(rango.min_row, rango.max_row + 1):
            for col_r in range(rango.min_col, rango.max_col + 1):
                valores_merge[(fila_r, col_r)] = valor_ancla

    def valor_celda(fila_idx, col_idx):
        if (fila_idx, col_idx) in valores_merge:
            return valores_merge[(fila_idx, col_idx)]
        return ws.cell(row=fila_idx, column=col_idx).value

    def partir_hes(hes_val):
        hes_raw = str(hes_val).strip() if hes_val is not None else ""
        for sep in (";", "\n"):
            hes_raw = hes_raw.replace(sep, ",")
        return [h.strip() for h in hes_raw.split(",") if h.strip()]

    grupos_bloque = []
    bloque_actual = None
    ocs_del_bloque = None  # dict: oc_str -> {"hes_list", "filas", "fila_principal", "mapa_hes_fila"}
    oc_actual = None       # referencia directa a la entrada del dict de arriba

    def cerrar_bloque_actual():
        nonlocal bloque_actual, ocs_del_bloque, oc_actual
        if bloque_actual is not None:
            ocs_list = [
                (oc, datos["hes_list"], datos["filas"], datos["fila_principal"], datos["mapa_hes_fila"])
                for oc, datos in (ocs_del_bloque or {}).items()
            ]
            grupos_bloque.append((bloque_actual, ocs_list))
        bloque_actual = None
        ocs_del_bloque = None
        oc_actual = None

    for fila_idx in range(fila_inicio, ws.max_row + 1):
        bloque_val = valor_celda(fila_idx, col_bloque)
        oc_val = valor_celda(fila_idx, col_oc)
        hes_val = valor_celda(fila_idx, col_hes)

        bloque_str = str(bloque_val).strip() if bloque_val not in (None, "") else ""
        oc_str = str(oc_val).strip() if oc_val not in (None, "") else ""
        hes_list_fila = partir_hes(hes_val)

        if not bloque_str and not oc_str and not hes_list_fila:
            continue

        if log:
            log(
                f"  [excel] fila {fila_idx}: Bloque_celda={bloque_str!r}  "
                f"OC_celda={oc_str!r}  HES_celda={hes_list_fila!r}"
            )

        if bloque_str and bloque_str != bloque_actual:
            cerrar_bloque_actual()
            bloque_actual = bloque_str
            ocs_del_bloque = {}
        elif bloque_actual is None:
            continue

        if oc_str:
            if oc_str in ocs_del_bloque:
                # La OC ya habia aparecido antes en este mismo bloque
                # (en una fila no contigua): se reengancha al mismo
                # grupo en vez de crear uno nuevo, para que sus HES se
                # sumen a la misma carpeta y no se reprocese la OC.
                oc_actual = ocs_del_bloque[oc_str]
                if log:
                    log(
                        f"  [excel]      OC {oc_str} repetida en el bloque "
                        f"{bloque_actual} (fila {fila_idx}): se agrupan sus "
                        f"HES con la aparicion anterior (fila principal "
                        f"{oc_actual['fila_principal']})."
                    )
            else:
                oc_actual = {
                    "hes_list": [],
                    "filas": [],
                    "fila_principal": fila_idx,
                    "mapa_hes_fila": {},
                }
                ocs_del_bloque[oc_str] = oc_actual
        elif oc_actual is None:
            continue

        oc_actual["hes_list"].extend(hes_list_fila)
        oc_actual["filas"].append(fila_idx)
        for h in hes_list_fila:
            if h not in oc_actual["mapa_hes_fila"]:
                oc_actual["mapa_hes_fila"][h] = fila_idx

    cerrar_bloque_actual()

    if log:
        for bloque, ocs in grupos_bloque:
            log(f"  [excel] => Bloque {bloque}: {len(ocs)} OC")
            for oc, hes_list, filas, fila_principal, _mapa_hes_fila in ocs:
                log(
                    f"  [excel]      OC {oc}: {len(hes_list)} HES agrupadas {hes_list!r} "
                    f"(filas Excel: {filas}, fila principal: {fila_principal})"
                )

    return grupos_bloque


# ---------------------------------------------------------------------------
# MEJORA 5: BACKUP DEL EXCEL Y MEJORA 8: VALIDACION DE EXCEL ABIERTO
# ---------------------------------------------------------------------------
def hacer_backup_excel(ruta_excel, log):
    """Guarda una copia del Excel de entrada, con marca de tiempo, en la
    misma carpeta donde esta el archivo, antes de empezar a escribir
    resultados en el. Es una red de seguridad barata por si algo saliera
    mal durante uno de los guardados de la corrida (corte de luz,
    archivo corrupto, etc). Falla en silencio salvo por un aviso en el
    log: no queremos que un backup fallido detenga toda la descarga."""
    try:
        carpeta = os.path.dirname(os.path.abspath(ruta_excel))
        nombre = os.path.basename(ruta_excel)
        base, ext = os.path.splitext(nombre)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_backup = os.path.join(carpeta, f"{base}_backup_{timestamp}{ext}")
        shutil.copy2(ruta_excel, ruta_backup)
        log(f"Backup del Excel original guardado en: {ruta_backup}")
    except Exception as e:
        log(f"Aviso: no se pudo crear el backup del Excel ({e}). Se continua igual.")


def excel_esta_bloqueado(ruta_excel):
    """True si el Excel parece estar abierto en otro programa (Excel
    suele mantener un lock exclusivo sobre el archivo mientras esta
    abierto). Se usa para avisar ANTES de arrancar la descarga, en vez
    de que el usuario se entere recien al fallar el primer guardado
    despues de la primera OC."""
    try:
        with open(ruta_excel, "r+b"):
            pass
        return False
    except PermissionError:
        return True
    except Exception:
        # Cualquier otro error (por ejemplo, permisos de carpeta) no lo
        # tratamos como "esta abierto"; si hay un problema real, va a
        # volver a fallar mas claramente al intentar leerlo/guardarlo.
        return False


# ---------------------------------------------------------------------------
# MEJORA 4: CONFIGURACION PERSISTENTE
# ---------------------------------------------------------------------------
def cargar_config():
    """Lee la configuracion guardada de una corrida anterior, si existe.
    Devuelve un dict vacio si no hay archivo de configuracion o si esta
    corrupto (nunca lanza excepcion, para no bloquear el arranque de la
    interfaz por esto)."""
    try:
        if os.path.isfile(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def guardar_config(valores):
    """Guarda la configuracion actual de la interfaz para la proxima
    corrida. Falla en silencio (no interrumpe el inicio de la descarga
    si por algun motivo no se puede escribir el archivo)."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(valores, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# PROCESO PRINCIPAL (corre en un hilo separado para no congelar la interfaz)
# ---------------------------------------------------------------------------
def procesar(
    ruta_excel,
    carpeta_destino,
    log,
    on_finish,
    nombre_hoja=NOMBRE_HOJA,
    col_bloque_letter=COL_BLOQUE_LETTER,
    col_oc_letter=COL_OC_LETTER,
    col_hes_letter=COL_HES_LETTER,
    fila_inicio=FILA_INICIO_DATOS,
    col_fecha_letter=COL_FECHA_LETTER,
    col_responsable_letter=COL_RESPONSABLE_LETTER,
    col_observaciones_letter=COL_OBSERVACIONES_LETTER,
    col_comentario_letter=COL_COMENTARIO_LETTER,
    fecha_valor="",
    responsable_valor="",
    pausa_entre_oc=PAUSA_ENTRE_OC,
    reanudar=True,
    detener_event=None,
    on_progress=None,
):
    """
    reanudar: si True (mejora 1), antes de tocar SAP para una OC se
        revisa si ella y todas sus HES ya tienen archivos descargados en
        disco; si es asi, se omite por completo (no se abre en SAP) y
        solo se recalculan Fecha/Responsable/Observaciones. Tambien se
        omiten individualmente las HES que ya tengan archivos dentro de
        una OC que no este completa. Si es False, se reprocesa todo
        desde cero como en la version original.
    detener_event: threading.Event opcional (mejora 2); si se activa
        mientras el proceso corre, se corta de forma ordenada entre OC,
        entre HES, o durante la pausa entre OC.
    on_progress: callback opcional on_progress(idx_global, total_oc, oc)
        que se llama despues de procesar (o de omitir) cada OC, para que
        la interfaz pueda mostrar progreso.
    """
    detenido_por_usuario = False
    try:
        # Backup del Excel original ANTES de empezar a tocarlo (mejora 5).
        hacer_backup_excel(ruta_excel, log)

        bloques = leer_filas_excel(
            ruta_excel,
            nombre_hoja=nombre_hoja,
            col_bloque_letter=col_bloque_letter,
            col_oc_letter=col_oc_letter,
            col_hes_letter=col_hes_letter,
            fila_inicio=fila_inicio,
            log=log,
        )
        total_oc = sum(len(ocs) for _, ocs in bloques)
        log(
            f"Se encontraron {len(bloques)} bloques con un total de {total_oc} OC "
            f"en el Excel (hoja '{nombre_hoja}')."
        )
        if reanudar:
            log("Reanudacion activada: las OC/HES que ya tengan archivos descargados se omitiran.")

        col_fecha_idx = openpyxl.utils.column_index_from_string(col_fecha_letter.strip().upper())
        col_responsable_idx = openpyxl.utils.column_index_from_string(col_responsable_letter.strip().upper())
        col_observaciones_idx = openpyxl.utils.column_index_from_string(col_observaciones_letter.strip().upper())
        col_comentario_idx = openpyxl.utils.column_index_from_string(col_comentario_letter.strip().upper())

        wb_write = openpyxl.load_workbook(ruta_excel)
        if nombre_hoja not in wb_write.sheetnames:
            raise ValueError(
                f"El Excel no tiene una hoja llamada '{nombre_hoja}'. "
                f"Hojas disponibles: {', '.join(wb_write.sheetnames)}"
            )
        ws_write = wb_write[nombre_hoja]

        session = get_sap_session()

        idx_global = 0
        for bloque, ocs in bloques:
            if detenido_por_usuario:
                break
            log(f"=== Bloque {bloque}: {len(ocs)} OC ===")
            carpeta_bloque = os.path.join(carpeta_destino, bloque)

            for oc, hes_list, filas_excel, fila_oc_principal, mapa_hes_fila in ocs:
                idx_global += 1

                # --- Mejora 2: revisar si se pidio detener el proceso ---
                if detener_event is not None and detener_event.is_set():
                    log("Proceso detenido por el usuario.")
                    detenido_por_usuario = True
                    break

                log(f"[{idx_global}/{total_oc}] Bloque {bloque} -> Procesando OC {oc}...")
                carpeta_oc = os.path.join(carpeta_bloque, oc)
                os.makedirs(carpeta_oc, exist_ok=True)

                # --- Mejora 1: omitir la OC completa si ya esta descargada ---
                if reanudar and _oc_esta_completa_en_disco(carpeta_oc, hes_list):
                    log(f"    (OC {oc} ya esta completa en disco -junto con todas sus HES-, se omite en SAP)")
                    hes_tiene_archivos = {h: True for h in hes_list}
                    hubo_comentario_oc = False
                    hes_tiene_comentario = {h: False for h in hes_list}
                else:
                    # --- Mejora 6: verificar que la sesion SAP siga viva ---
                    if not _sesion_esta_viva(session):
                        log("  Aviso: se perdio la conexion con la sesion de SAP, intentando reconectar...")
                        nueva_sesion = _reconectar_sesion(log)
                        if nueva_sesion is None:
                            log(
                                "ERROR GENERAL: no se pudo reconectar con SAP tras varios intentos. "
                                "Se detiene el proceso; los datos ya guardados en el Excel se conservan."
                            )
                            detenido_por_usuario = True
                            break
                        session = nueva_sesion

                    hubo_comentario_oc = False
                    if reanudar and _carpeta_tiene_archivos(carpeta_oc):
                        log(f"    (OC {oc} ya tiene archivos descargados, se omite la descarga a nivel OC)")
                    else:
                        try:
                            abrir_oc(session, oc, log)
                            if abrir_lista_anexos(session, log):
                                hubo_comentario_oc = descargar_todos_los_anexos(session, carpeta_oc, log)
                            else:
                                log(f"    (OC {oc} sin anexos, nada que descargar)")
                        except Exception as e:
                            log(f"  ERROR en OC {oc}: {e}")
                        finally:
                            _recuperar_sesion(session, log)
                            try:
                                volver_pantalla_inicial(session, log)
                                time.sleep(ESPERA)
                            except Exception as e:
                                log(f"  Aviso: no se pudo volver a pantalla inicial tras OC {oc}: {e}")
                                _recuperar_sesion(session, log)

                    if not hes_list:
                        log(f"    (OC {oc} no tiene HES asociadas en el Excel)")

                    hes_tiene_archivos = {}
                    hes_tiene_comentario = {}
                    for hes in hes_list:
                        if detener_event is not None and detener_event.is_set():
                            log("Proceso detenido por el usuario (a mitad de las HES de una OC).")
                            detenido_por_usuario = True
                            break

                        carpeta_hes = os.path.join(carpeta_oc, hes)
                        os.makedirs(carpeta_hes, exist_ok=True)

                        # --- Mejora 1: omitir la HES si ya esta descargada ---
                        if reanudar and _carpeta_tiene_archivos(carpeta_hes):
                            log(f"    -> HES {hes} ya tiene archivos descargados, se omite (reanudacion).")
                            hes_tiene_archivos[hes] = True
                            hes_tiene_comentario[hes] = False
                            continue

                        log(f"    -> Procesando HES {hes}...")
                        hubo_comentario_hes = False
                        try:
                            abrir_hes(session, hes, log)
                            if abrir_lista_anexos(session, log):
                                hubo_comentario_hes = descargar_todos_los_anexos(session, carpeta_hes, log)
                            else:
                                log(f"      (HES {hes} sin anexos, nada que descargar)")
                        except Exception as e:
                            log(f"    ERROR en HES {hes}: {e}")
                        finally:
                            _recuperar_sesion(session, log)
                            try:
                                volver_pantalla_inicial(session, log)
                                time.sleep(ESPERA)
                            except Exception as e:
                                log(f"    Aviso: no se pudo volver a pantalla inicial tras HES {hes}: {e}")
                                _recuperar_sesion(session, log)

                        hes_tiene_archivos[hes] = _carpeta_tiene_archivos(carpeta_hes)
                        hes_tiene_comentario[hes] = hubo_comentario_hes

                    if detenido_por_usuario:
                        # Igual guardamos lo que se alcanzo a procesar de
                        # esta OC antes de cortar, para no perder progreso.
                        pass

                # --- Escritura de resultados de esta OC en el Excel -------
                oc_tiene_archivos = _carpeta_tiene_archivos(carpeta_oc)
                observacion = _calcular_observacion(oc_tiene_archivos, hes_list, hes_tiene_archivos)
                log(f"    [OC {oc}] Observacion: {observacion}")

                for fila in filas_excel:
                    ws_write.cell(row=fila, column=col_fecha_idx).value = fecha_valor
                    ws_write.cell(row=fila, column=col_responsable_idx).value = responsable_valor
                    ws_write.cell(row=fila, column=col_observaciones_idx).value = observacion

                if hubo_comentario_oc and fila_oc_principal is not None:
                    ws_write.cell(row=fila_oc_principal, column=col_comentario_idx).value = (
                        f"esta oc {oc} tiene un comentario"
                    )
                    log(f"    [OC {oc}] tiene comentario -> anotado en fila {fila_oc_principal}")

                for hes, tiene_comentario in hes_tiene_comentario.items():
                    if not tiene_comentario:
                        continue
                    fila_hes = mapa_hes_fila.get(hes)
                    if fila_hes is None:
                        log(f"    Aviso: HES {hes} tuvo comentario pero no se encontro su fila en el Excel.")
                        continue
                    ws_write.cell(row=fila_hes, column=col_comentario_idx).value = (
                        f"esta HES {hes} tiene un comentario"
                    )
                    log(f"    [HES {hes}] tiene comentario -> anotado en fila {fila_hes}")

                try:
                    wb_write.save(ruta_excel)
                except Exception as e:
                    log(
                        f"    Aviso: no se pudo guardar el Excel tras la OC {oc} "
                        f"(¿esta abierto en otro programa?): {e}"
                    )

                if on_progress is not None:
                    try:
                        on_progress(idx_global, total_oc, oc)
                    except Exception:
                        pass

                if detenido_por_usuario:
                    break

                # --- Pausa entre OC ------------------
                if idx_global < total_oc:
                    _pausar_entre_oc(pausa_entre_oc, log, detener_event=detener_event)
                    if detener_event is not None and detener_event.is_set():
                        log("Proceso detenido por el usuario (durante la pausa entre OC).")
                        detenido_por_usuario = True
                        break

        if detenido_por_usuario:
            log("Proceso detenido por el usuario. El progreso ya realizado quedo guardado en el Excel.")
        else:
            log("Proceso finalizado.")
    except Exception as e:
        log(f"ERROR GENERAL: {e}")
    finally:
        on_finish(detenido_por_usuario)


# ---------------------------------------------------------------------------
# INTERFAZ GRAFICA
# ---------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Automatizacion SAP - Descarga de anexos OC / HES")
        root.geometry("720x820")

        # Mejora 2: evento para poder detener el proceso desde la interfaz.
        self.detener_event = threading.Event()
        # Mejora 3: archivo de log de la corrida actual (se abre en iniciar()).
        self.log_file_handle = None

        frm = tk.Frame(root, padx=10, pady=10)
        frm.pack(fill="x")

        tk.Label(frm, text="Archivo Excel:").grid(row=0, column=0, sticky="w")
        self.excel_var = tk.StringVar()
        tk.Entry(frm, textvariable=self.excel_var, width=55).grid(row=0, column=1, padx=5)
        tk.Button(frm, text="Buscar...", command=self.elegir_excel).grid(row=0, column=2)

        tk.Label(frm, text="Carpeta de descargas:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.carpeta_var = tk.StringVar()
        tk.Entry(frm, textvariable=self.carpeta_var, width=55).grid(row=1, column=1, padx=5, pady=(8, 0))
        tk.Button(frm, text="Buscar...", command=self.elegir_carpeta).grid(row=1, column=2, pady=(8, 0))

        tk.Label(frm, text="Nombre de hoja:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.hoja_var = tk.StringVar(value=NOMBRE_HOJA)
        tk.Entry(frm, textvariable=self.hoja_var, width=25).grid(row=2, column=1, sticky="w", padx=5, pady=(8, 0))

        cfg = tk.Frame(frm)
        cfg.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))

        tk.Label(cfg, text="Columna Bloque:").grid(row=0, column=0, sticky="w")
        self.col_bloque_var = tk.StringVar(value=COL_BLOQUE_LETTER)
        tk.Entry(cfg, textvariable=self.col_bloque_var, width=5).grid(row=0, column=1, padx=(4, 16))

        tk.Label(cfg, text="Columna OC:").grid(row=0, column=2, sticky="w")
        self.col_oc_var = tk.StringVar(value=COL_OC_LETTER)
        tk.Entry(cfg, textvariable=self.col_oc_var, width=5).grid(row=0, column=3, padx=(4, 16))

        tk.Label(cfg, text="Columna HES:").grid(row=0, column=4, sticky="w")
        self.col_hes_var = tk.StringVar(value=COL_HES_LETTER)
        tk.Entry(cfg, textvariable=self.col_hes_var, width=5).grid(row=0, column=5, padx=(4, 16))

        tk.Label(cfg, text="Fila de inicio:").grid(row=0, column=6, sticky="w")
        self.fila_inicio_var = tk.StringVar(value=str(FILA_INICIO_DATOS))
        tk.Entry(cfg, textvariable=self.fila_inicio_var, width=5).grid(row=0, column=7, padx=(4, 0))

        tk.Label(
            frm,
            text="Ej: si tu Bloque esta en la columna B, tu OC en la D, tu HES en"
            " la H, y los datos empiezan en la fila 2, deja los valores por"
            " defecto. Ajusta las letras/numero si el diseño del Excel cambia.",
            fg="gray",
            wraplength=660,
            justify="left",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(4, 0))

        cfg2 = tk.Frame(frm)
        cfg2.grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))

        tk.Label(cfg2, text="Columna Fecha:").grid(row=0, column=0, sticky="w")
        self.col_fecha_var = tk.StringVar(value=COL_FECHA_LETTER)
        tk.Entry(cfg2, textvariable=self.col_fecha_var, width=5).grid(row=0, column=1, padx=(4, 16))

        tk.Label(cfg2, text="Columna Responsable:").grid(row=0, column=2, sticky="w")
        self.col_responsable_var = tk.StringVar(value=COL_RESPONSABLE_LETTER)
        tk.Entry(cfg2, textvariable=self.col_responsable_var, width=5).grid(row=0, column=3, padx=(4, 16))

        tk.Label(cfg2, text="Columna Observaciones:").grid(row=0, column=4, sticky="w")
        self.col_observaciones_var = tk.StringVar(value=COL_OBSERVACIONES_LETTER)
        tk.Entry(cfg2, textvariable=self.col_observaciones_var, width=5).grid(row=0, column=5, padx=(4, 16))

        tk.Label(cfg2, text="Columna Comentario:").grid(row=0, column=6, sticky="w")
        self.col_comentario_var = tk.StringVar(value=COL_COMENTARIO_LETTER)
        tk.Entry(cfg2, textvariable=self.col_comentario_var, width=5).grid(row=0, column=7, padx=(4, 0))

        datos = tk.Frame(frm)
        datos.grid(row=6, column=0, columnspan=3, sticky="w", pady=(10, 0))

        tk.Label(datos, text="Fecha de descarga (DD/MM/AAAA):").grid(row=0, column=0, sticky="w")
        self.fecha_var = tk.StringVar()
        tk.Entry(datos, textvariable=self.fecha_var, width=12).grid(row=0, column=1, padx=(4, 16))

        tk.Label(datos, text="Responsable:").grid(row=0, column=2, sticky="w")
        self.responsable_var = tk.StringVar()
        tk.Entry(datos, textvariable=self.responsable_var, width=20).grid(row=0, column=3, padx=(4, 16))

        tk.Label(datos, text="Pausa entre OC (segundos):").grid(row=0, column=4, sticky="w")
        self.pausa_var = tk.StringVar(value=str(PAUSA_ENTRE_OC))
        tk.Entry(datos, textvariable=self.pausa_var, width=6).grid(row=0, column=5, padx=(4, 0))

        # --- Mejora 1: casilla de reanudacion ---
        opciones = tk.Frame(frm)
        opciones.grid(row=7, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self.reanudar_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opciones,
            text="Reanudar (omitir OC/HES que ya tengan archivos descargados)",
            variable=self.reanudar_var,
        ).grid(row=0, column=0, sticky="w")

        tk.Label(
            frm,
            text="La Fecha y el Responsable se escriben tal cual en TODAS las filas de"
            " TODAS las OC de esta corrida. La Observacion se calcula sola segun si la"
            " OC y cada HES tuvieron archivos. El Comentario se llena solo cuando se"
            " detecta un comentario/nota de SAP. La Pausa entre OC espera esa cantidad"
            " de segundos entre el fin de una OC y el inicio de la siguiente (no se"
            " aplica tras la ultima ni cuando una OC se omite por reanudacion).",
            fg="gray",
            wraplength=660,
            justify="left",
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(4, 0))

        # --- Botones Iniciar / Detener y etiqueta de progreso ---
        botones = tk.Frame(root)
        botones.pack(pady=10)

        self.btn_iniciar = tk.Button(
            botones, text="Iniciar descarga", command=self.iniciar, bg="#0a5", fg="white", width=18
        )
        self.btn_iniciar.grid(row=0, column=0, padx=5)

        self.btn_detener = tk.Button(
            botones, text="Detener", command=self.detener, bg="#a00", fg="white",
            width=12, state="disabled",
        )
        self.btn_detener.grid(row=0, column=1, padx=5)

        self.progreso_var = tk.StringVar(value="")
        tk.Label(root, textvariable=self.progreso_var, fg="#333").pack()

        self.log_box = scrolledtext.ScrolledText(root, width=84, height=20, state="disabled")
        self.log_box.pack(padx=10, pady=(0, 10), fill="both", expand=True)

        # Mejora 4: cargar configuracion guardada de una corrida anterior.
        self._cargar_config_en_interfaz()

    # -- Mejora 4: configuracion persistente -------------------------------
    def _cargar_config_en_interfaz(self):
        config = cargar_config()
        if not config:
            return
        for nombre_campo in CAMPOS_CONFIG_PERSISTENTE:
            if nombre_campo in config and hasattr(self, nombre_campo):
                try:
                    getattr(self, nombre_campo).set(config[nombre_campo])
                except Exception:
                    pass

    def _guardar_config_actual(self):
        valores = {}
        for nombre_campo in CAMPOS_CONFIG_PERSISTENTE:
            if hasattr(self, nombre_campo):
                try:
                    valores[nombre_campo] = getattr(self, nombre_campo).get()
                except Exception:
                    pass
        guardar_config(valores)

    def elegir_excel(self):
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if ruta:
            self.excel_var.set(ruta)

    def elegir_carpeta(self):
        ruta = filedialog.askdirectory()
        if ruta:
            self.carpeta_var.set(ruta)

    def log(self, mensaje):
        """Escribe el mensaje en el cuadro de texto de la interfaz y,
        si hay un archivo de log abierto para esta corrida (mejora 3),
        tambien lo agrega ahi con marca de hora."""
        def _append():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", mensaje + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

        self.root.after(0, _append)

        if self.log_file_handle is not None:
            try:
                hora = datetime.now().strftime("%H:%M:%S")
                self.log_file_handle.write(f"[{hora}] {mensaje}\n")
                self.log_file_handle.flush()
            except Exception:
                pass

    def actualizar_progreso(self, idx_global, total_oc, oc):
        """Callback de progreso (mejora 8): actualiza la etiqueta de la
        interfaz con el avance actual. Se llama desde el hilo de fondo,
        asi que la actualizacion real de la UI se agenda con root.after."""
        texto = f"Progreso: OC {idx_global} de {total_oc} (ultima procesada: {oc})"
        self.root.after(0, lambda: self.progreso_var.set(texto))

    def detener(self):
        """Mejora 2: pide que el proceso se detenga de forma ordenada en
        el proximo punto de control (entre OC, entre HES, o durante la
        pausa entre OC)."""
        if messagebox.askyesno(
            "Detener proceso",
            "¿Seguro que quieres detener el proceso? Se cortara de forma "
            "ordenada en cuanto termine el paso actual; lo ya descargado "
            "y guardado en el Excel no se pierde.",
        ):
            self.log(">>> Se solicito detener el proceso, terminando de forma ordenada...")
            self.detener_event.set()
            self.btn_detener.configure(state="disabled", text="Deteniendo...")

    def iniciar(self):
        ruta_excel = self.excel_var.get().strip()
        carpeta_destino = self.carpeta_var.get().strip()
        nombre_hoja = self.hoja_var.get().strip()

        if not ruta_excel or not os.path.isfile(ruta_excel):
            messagebox.showerror("Error", "Selecciona un archivo Excel valido.")
            return
        if not carpeta_destino:
            messagebox.showerror("Error", "Selecciona una carpeta de destino.")
            return
        if not nombre_hoja:
            messagebox.showerror("Error", "Ingresa el nombre de la hoja del Excel (ej: Plan de Descarga).")
            return

        # --- Mejora 8: validar que el Excel no este abierto en otro programa ---
        if excel_esta_bloqueado(ruta_excel):
            messagebox.showerror(
                "Error",
                "El archivo Excel parece estar abierto en otro programa (por "
                "ejemplo, Excel). Cierralo antes de iniciar la descarga, para "
                "que el script pueda guardar los resultados sin problemas.",
            )
            return

        col_bloque_letter = self.col_bloque_var.get().strip().upper()
        col_oc_letter = self.col_oc_var.get().strip().upper()
        col_hes_letter = self.col_hes_var.get().strip().upper()
        fila_inicio_str = self.fila_inicio_var.get().strip()

        col_fecha_letter = self.col_fecha_var.get().strip().upper()
        col_responsable_letter = self.col_responsable_var.get().strip().upper()
        col_observaciones_letter = self.col_observaciones_var.get().strip().upper()
        col_comentario_letter = self.col_comentario_var.get().strip().upper()
        fecha_valor = self.fecha_var.get().strip()
        responsable_valor = self.responsable_var.get().strip()
        pausa_str = self.pausa_var.get().strip()
        reanudar = bool(self.reanudar_var.get())

        for etiqueta, letra in (
            ("Bloque", col_bloque_letter),
            ("OC", col_oc_letter),
            ("HES", col_hes_letter),
            ("Fecha", col_fecha_letter),
            ("Responsable", col_responsable_letter),
            ("Observaciones", col_observaciones_letter),
            ("Comentario", col_comentario_letter),
        ):
            try:
                openpyxl.utils.column_index_from_string(letra)
            except Exception:
                messagebox.showerror(
                    "Error",
                    f"'{letra}' no es una columna de Excel valida para {etiqueta} (ej: B, D, H, Y, AA, AB).",
                )
                return

        try:
            fila_inicio = int(fila_inicio_str)
            if fila_inicio < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Error", "La fila de inicio debe ser un numero entero mayor o igual a 1."
            )
            return

        try:
            pausa_entre_oc = int(pausa_str)
            if pausa_entre_oc < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Error", "La Pausa entre OC debe ser un numero entero mayor o igual a 0 (segundos)."
            )
            return

        if not re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", fecha_valor):
            messagebox.showerror(
                "Error", "La Fecha de descarga debe tener el formato DD/MM/AAAA (ej: 30/08/2026)."
            )
            return

        if not responsable_valor:
            messagebox.showerror("Error", "Ingresa el nombre del Responsable.")
            return

        carpeta_destino = os.path.normpath(carpeta_destino)

        # --- Mejora 4: guardar la configuracion actual para la proxima vez ---
        self._guardar_config_actual()

        # --- Mejora 3: abrir archivo de log de esta corrida ---
        self.log_file_handle = None
        try:
            carpeta_logs = os.path.join(carpeta_destino, "logs")
            os.makedirs(carpeta_logs, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            ruta_log = os.path.join(carpeta_logs, f"log_{timestamp}.txt")
            self.log_file_handle = open(ruta_log, "w", encoding="utf-8")
        except Exception as e:
            # No bloqueamos la descarga por esto, solo avisamos que el log
            # de esta corrida no va a quedar guardado en archivo.
            messagebox.showwarning(
                "Aviso",
                f"No se pudo crear el archivo de log de esta corrida ({e}). "
                "El proceso va a continuar igual, solo que el log no quedara "
                "guardado en disco.",
            )

        self.detener_event.clear()
        self.progreso_var.set("Progreso: iniciando...")
        self.btn_iniciar.configure(state="disabled", text="Procesando...")
        self.btn_detener.configure(state="normal", text="Detener")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        hilo = threading.Thread(
            target=procesar,
            args=(ruta_excel, carpeta_destino, self.log, self.on_finish),
            kwargs=dict(
                nombre_hoja=nombre_hoja,
                col_bloque_letter=col_bloque_letter,
                col_oc_letter=col_oc_letter,
                col_hes_letter=col_hes_letter,
                fila_inicio=fila_inicio,
                col_fecha_letter=col_fecha_letter,
                col_responsable_letter=col_responsable_letter,
                col_observaciones_letter=col_observaciones_letter,
                col_comentario_letter=col_comentario_letter,
                fecha_valor=fecha_valor,
                responsable_valor=responsable_valor,
                pausa_entre_oc=pausa_entre_oc,
                reanudar=reanudar,
                detener_event=self.detener_event,
                on_progress=self.actualizar_progreso,
            ),
            daemon=True,
        )
        hilo.start()

    def on_finish(self, detenido_por_usuario=False):
        def _finalizar():
            self.btn_iniciar.configure(state="normal", text="Iniciar descarga")
            self.btn_detener.configure(state="disabled", text="Detener")

            if self.log_file_handle is not None:
                try:
                    self.log_file_handle.close()
                except Exception:
                    pass
                self.log_file_handle = None

            # --- Mejora 7: notificacion al terminar ---
            try:
                import winsound
                winsound.MessageBeep(
                    winsound.MB_ICONEXCLAMATION if detenido_por_usuario else winsound.MB_ICONASTERISK
                )
            except Exception:
                pass  # no estamos en Windows, o el sonido no esta disponible; no es critico

            if detenido_por_usuario:
                messagebox.showinfo(
                    "Proceso detenido",
                    "El proceso se detuvo por pedido del usuario. Lo ya "
                    "descargado y guardado en el Excel se conservo; puedes "
                    "volver a iniciar la descarga (con 'Reanudar' activado) "
                    "para continuar donde quedo.",
                )
            else:
                messagebox.showinfo(
                    "Proceso finalizado",
                    "La descarga de anexos termino. Revisa el log para ver "
                    "el detalle de cada OC/HES procesada.",
                )

        self.root.after(0, _finalizar)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
