# Optimizacion y fluidez

## Cambios

- Busqueda, completado y drivers ejecutan captura, procesamiento y exportacion
  en un worker. El hilo de Tkinter recoge eventos mediante una cola y `after`.
  Los widgets y dialogos se actualizan solamente desde el hilo de interfaz.
- Un coordinador compartido por ventana admite un trabajo a la vez. El boton
  activo se deshabilita y muestra la etapa. Los errores restauran el boton y
  permiten reintentar. La configuracion se copia al iniciar: editar controles
  no modifica la ejecucion en curso.
- Las etapas son captura, validacion de encabezados, procesamiento y guardado.
  El log tecnico registra duracion por etapa y total. No se muestran porcentajes
  estimados. La etapa de procesamiento incluye la lectura de snapshots.
- Si se intenta cerrar durante el procesamiento, la ventana permanece abierta
  y pide esperar al final; no interrumpe el guardado. No hay cancelacion forzada.
- Similaridad reutiliza el mejor match de cada consulta normalizada repetida.
  La cache es local a una ejecucion y a su base: nunca comparte resultados entre
  archivos, runs, umbrales o snapshots. Se conserva el scorer, las variantes,
  orden de candidatos, desempates, puntajes y resultados por cada fila.
- Drivers decide formatos una vez por columna, los aplica al escribir y calcula
  anchos sobre las primeras 200 filas. Evita recorrer de nuevo todas las celdas
  y consultar el encabezado para cada una. Mantiene los formatos y anchos.
- Se conserva `iterrows()` en las reglas de drivers. Estas reglas involucran
  valores heterogeneos y prioridades; no se sustituyeron sin evidencia de que
  sean el cuello de botella. Tampoco se redujeron las validaciones de integridad.

## Medicion reproducible

```powershell
.\.venv\Scripts\python.exe -m scripts.benchmark_optimization --repeats 3
```

El script carga dos funciones originales del commit `55747c5` y compara sus
resultados con las actuales. Usa datos sinteticos deterministas, mediana de
tres ejecuciones, Windows y Python 3.13.0. Verifica igualdad de filas de match,
valores de celdas, estilos, anchos, filtros y paneles congelados. No toca Excel
de usuario ni almacena resultados de matching en el repositorio.

Medicion realizada el 2026-10-05:

| Escenario | Antes | Despues | Mejora |
|---|---:|---:|---:|
| Similaridad: 700 filas / base 6000, 700 consultas distintas | 2.9474 s | 2.9189 s | Practicamente igual |
| Similaridad: 700 filas / base 6000, 50 consultas distintas repetidas | 2.9226 s | 0.2364 s | 12.37 veces |
| Construccion de hoja driver: 6000 filas / 8 columnas | 0.2026 s | 0.1730 s | 14.6% menos tiempo |

Estas medidas aislan el algoritmo de similaridad y la construccion/formato de
una hoja. No representan el tiempo total de abrir, capturar, comprimir y guardar
Excel. La mejora depende de cuantos valores se repitan y del contenido real.
Los threads mantienen operativa la interfaz; no prometen acelerar tareas CPU
ni ejecutar varias busquedas en paralelo.

## Tests

`test_optimization.py` verifica consultas repetidas, desempates, puntajes,
cache renovada tras cambiar el maestro, completado por fila, formatos, anchos
y eventos de progreso. `test_background_tasks.py` comprueba el worker separado,
callbacks en el hilo Tk, heartbeat de interfaz durante procesamiento detenido,
exclusion de una segunda ejecucion, errores y reintento, configuracion congelada,
las tres pipelines con Excel reales y cierre protegido. La suite de integridad
anterior se ejecuta junto con estas pruebas.

Para QAS basta una ejecucion habitual: mover la ventana o cambiar de pestaña
durante el procesamiento, observar las etapas y confirmar el resultado conocido.
Las mediciones de archivos empresariales muy grandes y la sesion real de
Excel/OneDrive siguen requiriendo datos representativos.
