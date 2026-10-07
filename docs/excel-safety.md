# Manejo seguro de Excel

## Baseline y comportamiento preservado

CP0 se realizo sobre `55747c5`. No habia tests versionados. Antes de modificar
produccion se ejecutaron nueve casos de regresion con OpenPyXL/Pandas reales.
Pasaron con las versiones de `requirements.txt`, Python 3.13.0 y Windows.
El README original indicaba Python 3.12; esa version no se ha verificado aqui.

Se mantienen la normalizacion, primera coincidencia exacta entre duplicados,
similaridad y umbrales, columnas tecnicas, mapeos de completado, prioridades
CECO, expansion por DNI, calculos y formato de salida. Los cuerpos de los
procesadores quedaron en `_run_search`, `_run_completion` y `_run_driver`.
Los metodos publicos mantienen sus argumentos y retornos y agregan captura.

La configuracion consulta hojas/encabezados/muestras y cierra el workbook.
Los perfiles mantienen su formato. Los datos pueden cambiar despues de
configurar: se toman los existentes al comenzar la ejecucion. Las columnas
configuradas se verifican de nuevo en la captura; un cambio de encabezados
requiere releerlas. Los ejecutores drivers y completado tambien verifican
encabezados completos cuando la GUI proporciona esa metadata. Los perfiles
anteriores siguen funcionando.

## Checkpoints implementados

| CP | Objetivo y archivos | Riesgo y verificacion |
|---|---|---|
| 1 | `tests/`, `pytest.ini`, `requirements-dev.txt`: fijar resultados actuales | Nueve casos pasaron antes de tocar produccion. Sin cambios de negocio. |
| 2 | `excel_config_panel.py`, `excel_inspector.py`: consultar y cerrar | Reapertura en cada consulta; tests con panel Tk real y acceso exclusivo Windows despues de consultas y errores. |
| 3 | `excel_safety.py`, ejecutores, metadata de paneles: verificar encabezados al ejecutar | Datos editados se aceptan; encabezados configurados cambiados se rechazan. |
| 4 | Captura conjunta en `excel_safety.py` y wrappers de ejecutores | SHA256 de todas las fuentes antes, copias y fuentes despues. Se reintenta el conjunto si cambia cualquier fuente. Tests de cambios durante y despues de captura. |
| 5 | Exportadores y `atomic_write`/`atomic_update`: persistir temporal validado | ZIP con CRC, estructura de workbook, apertura OpenPyXL, flush/fsync antes de publicar. Tests de guardado parcial, ZIP invalido y fallo de replace. |
| 6 | Guard de destino, hash esperado y señales Office/Windows | Captura esperada en `DataFrame.attrs`; chequeos antes de construir y antes de publicar. Pruebas de cambios, eliminacion, segunda instancia y bloqueo Windows sin ~$ . |
| 7 | `tests/`, `app.py`, `.gitignore`, esta documentacion | Crash real de subprocess, reintentos limitados, limpieza, macros sinteticas, tablas, nombres, formulas, hoja oculta y GUI. |

`git diff` muestra modificaciones de produccion/documentacion. Los archivos
nuevos de infraestructura y tests se ven con `git status --short`; un diff sin
staging no incluye archivos nuevos. Los checkpoints se validaron localmente
antes de publicar la rama de trabajo para revision mediante pull request.

## Captura y maestro de solo lectura

Se eligio captura **del conjunto** en cada ejecucion. Las entradas forman el
mismo cruce y no deben cambiar durante su captura. Esto no demuestra que
ambos libros correspondan a una misma version empresarial: esa relacion
necesita una identificacion de negocio que hoy no existe en el programa.

La captura guarda rutas, SHA256, tamaño, mtime en nanosegundos, fecha UTC e
intento. Las copias estan en un TemporaryDirectory del sistema y se eliminan
al terminar, incluso cuando falla el procesamiento. La metadata permanece en
el DataFrame y en logs; las rutas snapshot son historicas, no archivos permanentes.
Los nombres/rutas originales del perfil no se sustituyen y `FUENTE_ARCHIVO`
continua mostrando el nombre original del maestro.

El maestro se lee y no se escribe. Si se selecciona el mismo archivo como
maestro y destino, la actualizacion del original se rechaza; puede exportarse
a un archivo nuevo. No se rechaza una entrada legible solo por existir ~$:
esa señal se utiliza para impedir escritura, no para impedir leer un snapshot.

## Publicacion

En actualizaciones, el resultado conserva el hash del destino observado al
capturar. El exportador exige esta metadata: un DataFrame construido por otro
medio no puede actualizar el original sin pasar por el procesamiento.

Se coordina el destino por ruta resuelta usando un lock de sistema operativo.
El lock se libera con el cierre del proceso, tambien si este muere. El archivo
de lock queda en el temporal del sistema y no se borra para evitar carreras
entre instancias que ya lo abrieron. La publicación tiene cinco intentos con esperas de 0.5, 1, 1.5 y 2 segundos
para bloqueos transitorios (WinError 5, 32 y 33). Cada intento verifica de nuevo
el destino. Cambios externos y archivos inválidos no se reintentan al publicar.
Las otras operaciones conservan los tres intentos de 0.15 y 0.30 segundos.

El guardado ocurre en un temporal hermano. Un writer cerrado produce el
temporal y se valida su integridad antes de publicar. La actualizacion usa
`os.replace`; un archivo nuevo usa `os.rename` en Windows (no reemplaza un
destino existente) o enlace no destructivo en otros sistemas. Un sufijo
aleatorio evita colisiones entre archivos nuevos generados en el mismo segundo.

Busqueda, completado y drivers conservan las hojas existentes y buscan un
nombre disponible: `RESULTADO`, `RESULTADO_2`, `RESULTADO_3`, etc. Se consideran
colisiones sin distinguir mayusculas/minusculas y despues de limpiar caracteres
invalidos/truncar a 31 caracteres. Drivers conserva tambien sus cinco hojas
de salidas anteriores y crea nuevas hojas con sufijo. La app muestra los
nombres finales en el log y el mensaje de resultado. Esta proteccion cambia
deliberadamente la politica destructiva inicial, sin alterar el matching.

Los libros `.xlsm` se abren con `keep_vba=True`. Se verifica conservacion de
los bytes VBA mediante un paquete sintetico; no se han ejecutado macros
reales ni validado firmas digitales. El archivo nuevo sigue siendo `.xlsx`.

## Diagnostico y limites

La app informa la ruta del diagnostico al iniciar. Los logs se guardan fuera
del repositorio, en `%TEMP%/presupuesto_excel_logs/excel_<pid>.log`, con rotacion
a 1 MB y dos copias por proceso. Incluyen hashes, captura, intentos, destino
y temporales que no se pudieron limpiar. La GUI muestra mensajes especificos
para archivo ocupado, cambio externo o encabezados cambiados.

Las garantias comprobadas son: originales intactos al fallar antes de publicar,
maestro intacto, temporales limpiados en salidas normales/errores, deteccion de
cambios observados y exclusion entre instancias del mismo usuario/equipo.

Los hashes y `os.replace` no forman un CAS indivisible frente a escritores
externos. Un proceso ajeno puede escribir entre el ultimo chequeo y el replace.
Los hashes tampoco detectan un archivo que cambia y vuelve exactamente a
sus bytes anteriores entre chequeos. No existe transaccion distribuida con
Excel, OneDrive, SharePoint ni instancias en otros equipos/usuarios. Una ruta
alternativa al mismo archivo puede no compartir el mismo lock.

Un kill impide ejecutar finally: puede quedar un temporal hermano incompleto
o un directorio de captura huerfano. Nunca se publica automáticamente ese temporal. Si ya se había preparado una
copia de recuperación válida, se puede recuperar desde Resultados pendientes. El test de crash confirma que el original permanece intacto y el
lock queda libre. Se pueden eliminar estos temporales despues de cerrar las
instancias; no se hace una limpieza indiscriminada de archivos.

fsync sincroniza el temporal; no se promete durabilidad absoluta ante corte
electrico, fallos de disco o implementaciones particulares de red. Se probaron
locks nativos Windows, no una sesion real de Microsoft Excel ni OneDrive.
OpenPyXL puede no preservar extensiones Excel no soportadas; tablas, nombres,
formulas y hojas ocultas comunes tienen tests, libros empresariales complejos
necesitan archivos representativos. El costo de captura/hash es lineal en el
tamaño de entradas; no se ha medido rendimiento con Excel de produccion.

El reporte opcional de debug drivers conserva su lectura existente; su salida
usa publicacion segura. Sus lecturas no constituyen el resultado del match.

## Ejecutar la verificacion

```powershell
.\.venv\Scripts\python.exe -m pytest -q -W error
git diff --check
git diff
git status --short
```

Los Excel de pruebas viven en `tmp_path`. No se versionan archivos de usuario,
perfiles locales, entornos virtuales ni resultados generados.


## Recuperación local de guardados

Antes de publicar un libro ya serializado y validado, se conserva su copia en
`%LOCALAPPDATA%/AppDriversPresupuestoTI/recovery` (temporal del sistema como
alternativa cuando LOCALAPPDATA no está definido). Los registros JSON y su
copia se sincronizan a disco; cada evento se publica por renombrado local.
No se ejecuta otra vez el match ni la generación de drivers al recuperar.

Resultados pendientes permite reintentar el destino original con su hash
esperado, o guardar una copia completa del libro en otro nombre inexistente.
Conserva la extensión XLSX/XLSM. No sobrescribe otro archivo elegido mediante
el diálogo. Los pendientes no bloquean otros trabajos ni ejecuciones futuras.
La copia está ligada al equipo/usuario; no viaja automáticamente con el ZIP
portable y puede contener todos los datos del libro original.

Tras guardar, se registra un recibo y se elimina el Excel de recuperación; se
conservan los pequeños recibos para reconocer un reintento ya completado. Si
el proceso termina entre la publicación y el recibo, la recuperación compara
el hash completo de los destinos intentados con la copia y evita repetir el
guardado si ya coincide. Si un tercero editó esa salida después del crash,
no se puede confirmar automáticamente su publicación; se conserva el pendiente.

Si el cálculo o la serialización falla antes de producir un Excel válido, no
hay libro recuperable. Si el archivo original no puede leerse para construir
la salida, tampoco se puede reconstruir su contenido completo. Un fallo del
almacenamiento local puede impedir conservar la copia y se comunica como error.
No existe garantía absoluta ante fallos de disco/corte eléctrico.

El guardado confirma únicamente el archivo local, nunca la sincronización de
OneDrive/SharePoint. No se ha realizado una prueba con OneDrive corporativo
real. Las pruebas simulan los códigos Windows en el reemplazo final, persistencia,
conflictos humanos, copia alternativa, crash entre publicación y recibo,
recuperación desde registros de disco y el panel GUI.
