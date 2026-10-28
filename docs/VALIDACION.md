# Validación de la preparación

Verificación realizada el 28 de octubre de 2026 con Python 3.12.14 en Linux.

| Verificación | Resultado |
| --- | --- |
| Suite de la aplicación | 42 pruebas aprobadas |
| Importación y ayuda de los scripts estadísticos | 2 pruebas aprobadas, con los 4 scripts verificados |
| Compilación Python de aplicación y análisis | Sin errores de sintaxis |
| Comparación de `app.py`, `iniciar.py`, `cfos_core.py` y `batch_project.py` contra las fuentes | Idénticos byte a byte |
| Plantillas de distribución | Encabezados válidos y cero animales |

Las versiones directas utilizadas se registran en `requirements-tested.txt`. Ese archivo permite instalar esas versiones, pero no bloquea todas las dependencias transitivas ni garantiza compatibilidad con cualquier sistema operativo.

La suite de la aplicación usa entradas temporales y datos sintéticos; algunas lecturas de ROI se sustituyen por polígonos de prueba. Se verifica una exportación del pipeline con esos datos. No se ha realizado una revisión visual interactiva de la GUI ni una corrida estadística completa con la cohorte experimental final durante esta preparación. El lanzador macOS se ha corregido por inspección; no se ha ejecutado en macOS.

Los catálogos originales y sus tamaños de cohorte no se incorporaron al repositorio. Sus pruebas se adaptaron para validar plantillas vacías y un catálogo sintético, manteniendo las demás comprobaciones.
