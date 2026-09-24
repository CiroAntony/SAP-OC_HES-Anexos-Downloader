# SAP-OC_HES-Anexos-Downloader

Descarga automática de anexos de **OC (órdenes de compra)** y **HES (hojas de entrada/servicios)** desde **SAP GUI**, con interfaz gráfica (Tkinter) en Windows.

El programa navega por las transacciones `ME23N` / `ML81N`, abre la lista de anexos de cada documento y los exporta a carpetas locales, escribiendo el resultado de vuelta en el Excel de entrada.

## Requisitos

- **Windows** con **Python 3.10 o superior**.
- **SAP Logon** abierto con una sesión **ya conectada** antes de ejecutar el script.
- **SAP GUI Scripting** habilitado (cliente y servidor).
- La opción **"Visualizar ventana de diálogo nativa de Microsoft Windows"** activada (Opciones de SAP GUI > Accesibilidad y Scripting > Scripting), necesaria para el diálogo nativo de "Exportar fichero" que maneja pywinauto.

## Instalación

```bash
py -m pip install pywin32 pywinauto openpyxl
```

Tkinter ya viene incluido con Python en Windows.

## Uso

1. Abre SAP Logon e inicia sesión.
2. Ejecuta `descarga_sap.py` desde la carpeta del proyecto:

   ```bash
   py descarga_sap.py
   ```

3. Conserva `sap_backend.py` junto a `descarga_sap.py`: es el backend de automatización que usa el programa.

### Fuente de datos

- **Desde Excel:** elige un archivo `.xlsx` o `.xlsm`, la hoja, la fila inicial y las letras de columna de Bloque, OC y HES. La columna de Bloque puede dejarse vacía si no existe. Puedes procesar solo OC, solo HES o ambos. Las celdas combinadas se leen por su valor visible y los códigos HES pueden separarse con coma, punto y coma o salto de línea.
- **Manual:** escribe un código por línea. Para combinar tipos usa `OC 4500001234` y `HES 10001234`.

### Estructura del Excel por defecto

| Columna | Contenido | Por defecto |
| --- | --- | --- |
| Bloque | Código de bloque | B |
| OC | Código de orden de compra | D |
| HES | Código(s) de hoja de entrada/servicios | H |
| Fila de inicio de datos | — | 2 |

La hoja usada por defecto es **"Plan de Descarga"**. El nombre de la hoja y las columnas son configurables desde la interfaz.

### Carpetas generadas

```
<carpeta_destino>/
  <Bloque>/
    OC/<codigo_oc>/        -> anexos de la propia OC
    OC/<codigo_oc>/HES/<codigo_hes>/   -> anexos de cada HES de esa OC
```

`reporte_descargas.csv` se actualiza tras cada documento con fecha, bloque, tipo, código, filas de origen, estado, error y carpeta. El Excel de entrada no se modifica (salvo la escritura de resultados descrita abajo).

## Flujo de trabajo

1. Elige la fuente de datos y configura las columnas.
2. Revisa la **vista previa** con los documentos detectados antes de iniciar. Los códigos repetidos se descargan una vez por tipo y bloque; en modo OC + HES, las HES se guardan bajo su OC.
3. Inicia la descarga. El programa se conecta a la sesión SAP activa y procesa cada documento.

### Opciones

- **Reanudar:** omite los documentos cuya carpeta ya tenga archivos de tamaño mayor que cero. Desactívala si necesitas volver a descargar un documento concreto.
- **Detener:** detiene el proceso de forma ordenada al terminar el documento en curso.
- **Pausa (segundos):** tiempo de espera entre documentos para simular un ritmo manual.

## Notas técnicas

- **Diálogo "Exportar fichero":** se maneja con pywinauto usando `backend="win32"` (el backend `uia` no detecta este diálogo en SAP GUI 760). La llamada `pressToolbarButton("%ATTA_EXPORT")` no retorna hasta que se cierra el diálogo, por lo que se dispara en un hilo aparte para evitar un *deadlock*.
- **Comentarios SAP:** si una fila de la lista de anexos es un comentario/nota (no un archivo exportable), se detecta y se anota.
- **Archivos pesados:** la espera por la exportación es adaptativa (monitorea `session.Busy` y la estabilización del archivo en disco) y los errores transitorios por "SAP ocupado" se reintentan automáticamente.
- **Escritura de resultados en el Excel (solo en `sap_backend.py`, versión completa):** las columnas de Fecha, Responsable, Observaciones y Comentario (Y, Z, AA, AB por defecto) se rellenan en el propio Excel, con backup previo y validación de que el archivo no esté abierto.
- **Recuperación de sesión:** se cierran ventanas colgadas tras cada documento y se verifica/reconecta la sesión SAP si la conexión se pierde.

## Alcance de verificación

La navegación en ME23N/ML81N y la exportación usan la lógica de scripting original; no hay sesión SAP disponible en un entorno de desarrollo para probarlas. Conviene empezar con una OC y una HES de prueba y revisar el CSV y las carpetas generadas.