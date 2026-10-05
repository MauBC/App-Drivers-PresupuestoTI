# App Drivers - Presupuesto TI

Aplicación de escritorio desarrollada en Python para automatización y procesamiento de archivos Excel relacionados con presupuesto y drivers.

## Tecnologías

- Python 3.12
- CustomTkinter
- pandas
- openpyxl
- RapidFuzz
- XlsxWriter
- PyInstaller

## Instalación

```powershell
pip install -r requirements.txt
```

## Pruebas automaticas

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest -q -W error
```

Las pruebas crean Excel pequeños en directorios temporales. Cubren busqueda,
completado, drivers, los paneles de configuracion, bloqueos de Windows,
cambios concurrentes y fallos de guardado. Las pruebas de GUI necesitan una
sesion de escritorio con Tk disponible; no abren cuadros de dialogo.

El diseño, checkpoints y limites estan documentados en
[docs/excel-safety.md](docs/excel-safety.md).

La ejecucion en segundo plano y las mediciones de rendimiento estan en
[docs/optimization.md](docs/optimization.md).

Las salidas dentro de un libro conservan todas sus hojas anteriores. Si el
nombre solicitado existe, se genera uno libre con sufijo (`RESULTADO_2`,
`DRIVER_2`, etc.). El log y el mensaje final indican las hojas creadas.
