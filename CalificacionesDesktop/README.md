# Johana Software

Aplicación WPF para gestión de calificaciones.

## Cambios de esta versión – Notas generales

- Las notas aplicadas se conservan en memoria durante toda la sesión, aunque se cambie de trimestre y se vuelva a analizar.
- Al guardar, se escriben todas las notas aplicadas durante la sesión (por ejemplo, Biología del primer trimestre e Inglés/Biología del segundo) en un solo archivo de salida.
- La comprobación de notas existentes se hace en la columna correspondiente a la asignatura y al trimestre seleccionados.
- Después de aplicar notas se muestra `Notas aplicadas correctamente.`
- Después de guardar correctamente, el único mensaje de confirmación es `Archivo descargado con éxito`.
- Al volver al menú principal y entrar nuevamente al módulo, se limpia la sesión y se comienza desde cero.
- Para conservar los elementos originales del libro, el guardado copia el archivo y escribe únicamente las celdas de notas mediante Microsoft Excel de escritorio.

Requisitos: .NET 8 para Windows, Visual Studio 2022 y Microsoft Excel de escritorio.
