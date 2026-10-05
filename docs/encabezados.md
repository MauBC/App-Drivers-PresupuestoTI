# Encabezados de Excel

La inspección, Drivers y Completar comparten reglas para los nombres de columnas.
Se recortan espacios externos (incluido el espacio no separable) y se eliminan
los caracteres invisibles U+200B, U+FEFF y U+2060. Se normaliza Unicode a NFC,
conservando tildes, ñ, símbolos y espacios internos. `Año` y `Ano` siguen siendo
columnas diferentes. Las referencias por nombre ignoran mayúsculas/minúsculas.

Dos encabezados que resulten equivalentes se rechazan con sus posiciones antes
de procesar datos, incluso si pandas intentara renombrarlos automáticamente.
Las celdas vacías reciben `COL_n`, donde n es su posición; también se rechazan
colisiones entre esos identificadores y encabezados existentes.

Estas reglas se aplican en memoria y no modifican encabezados de archivos de
entrada. Las nuevas hojas de salida usan los nombres limpios. En Completar,
los nombres técnicos se escriben con las mayúsculas que utiliza la aplicación.
La comprobación de encabezados configurados sigue detectando cambios reales;
los espacios externos y los caracteres invisibles admitidos no invalidan la
configuración por sí solos.

Pruebas: `python -m pytest -q`. Los casos de encabezados están en
`tests/test_excel_headers.py` e incluyen cruces, completar y distribución de
drivers con Excel reales y verificación de que las entradas no se modifican.
