# Heatmap cFos

Código de **Heatmap V1.2, versión interna 1.2.4**, para cuantificar objetos c-Fos y estudiar su distribución espacial en BLAa y BLAp. Incluye la aplicación de escritorio y los scripts del análisis final.

El programa trabaja con objetos previamente segmentados y exportados desde Imaris, archivos de área/volumen y ROI de ImageJ. No realiza la segmentación de imágenes de microscopía.

## Funciones

- Registro por experimento y animal, selección de animales y control de calidad por stack.
- Conteo de objetos antes del filtro, después del filtro y dentro de la ROI.
- Densidad por volumen, con el animal como unidad del resumen principal.
- Mapas 2D y comparación de la distribución medial–lateral y ventral–dorsal.
- Coordenadas 3D con Z en µm y posición anteroposterior en mm de bregma.
- Edición de bregma, actualización de experimentos y consolidación de Stack_ID duplicados.
- Exportación de tablas, matrices y figuras para el análisis posterior.

## Instalación y ejecución

Se ha comprobado el código con Python 3.12. La interfaz requiere Tkinter y una sesión gráfica.

Desde la carpeta del repositorio:

```bash
python -m venv .venv
```

Activa el entorno con `source .venv/bin/activate` en macOS/Linux o con `.venv\Scripts\Activate.ps1` en PowerShell. Después:

```bash
python -m pip install -r requirements.txt
python Heatmap_V1_2/iniciar.py
```

`python -m tkinter` permite comprobar Tkinter. En distribuciones Linux puede ser necesario instalar el paquete de Tkinter correspondiente al Python utilizado. En macOS también se conserva el lanzador `Heatmap_V1_2/INICIAR_HEATMAP_V1_2.command` para instalaciones con un entorno Conda llamado `cfos`.

## Primer uso

1. Abre la aplicación e importa tus catálogos con **Actualizar grupos** y **Actualizar bregmas (catálogo y proyecto)**. Las plantillas incluidas tienen únicamente encabezados.
2. Selecciona los archivos de objetos, área/volumen y ROI de cada experimento. Revisa la correspondencia por stack.
3. Comprueba el montaje, la orientación de la ROI, el bregma y el control de calidad. Guarda o actualiza el experimento.
4. Marca los animales que entrarán en el análisis y ejecuta la corrida.

El proyecto se guarda por defecto en `proyecto_cfos/`, junto a la carpeta `Heatmap_V1_2/`. También puedes seleccionar un proyecto existente desde la interfaz. La variable de entorno `HEATMAP_CFOS_PROJECT` permite indicar otra ubicación.

Los datos experimentales, catálogos completados y resultados se gestionan localmente. `.gitignore` excluye esas rutas y los formatos de datos; solo permite las dos plantillas vacías que necesita el arranque. No reemplaces esas plantillas versionadas con catálogos reales: impórtalos desde la interfaz.

## Convenciones anatómicas

| Variable | Convención |
| --- | --- |
| X normalizada | 0 = medial; 1 = lateral |
| Y normalizada | 0 = ventral; 1 = dorsal |
| Z | Profundidad real en µm |
| AP | Posición de bregma en mm |
| Unidad principal | Animal; suma de objetos válidos / suma de volumen |

La transformación de Y es `Y_DV = 1 - Y_imagen`. Para el esquema de montaje de este estudio, el montaje izquierdo invierte X y el derecho invierte Z; la corrección de montaje no introduce otra inversión de Y. Estas convenciones requieren verificar cómo se adquirieron y montaron las muestras antes de aplicarlas a otro estudio.

El filtro elimina un objeto solo cuando se cumplen simultáneamente `Volume < 50 µm³` y `Surface / Volume > 2.8`. Consulta [docs/METODO.md](docs/METODO.md) para las unidades, entradas y límites de interpretación.

## Código de análisis estadístico

La carpeta [analysis/](analysis/README.md) contiene los scripts del análisis final, con rutas configurables por argumentos. Conserva el diseño del estudio: BLAa/BLAp y Vehículo/SCOP-HB/LiCl. Para ejecutarlos instala también:

```bash
python -m pip install -r requirements-analysis.txt
```

Sus instrucciones detallan el orden de ejecución, los archivos necesarios y los supuestos. El repositorio no contiene informes ni resultados del estudio.

## Pruebas

```bash
python -m unittest discover -s Heatmap_V1_2 -p "test_*.py" -v
python -m unittest discover -s tests -p "test_*.py" -v
```

Las pruebas de la aplicación cubren filtros, orientación, reconciliación de volúmenes por sector, exportación de coordenadas, selección de animales, bregma, duplicados y persistencia. Las de análisis comprueban la importación sin ejecutar el estudio y las interfaces de línea de comandos. Consulta [docs/VALIDACION.md](docs/VALIDACION.md) para el alcance de la verificación.

## Procedencia y derechos

Se conserva el aviso original: **Propiedad del Laboratorio de neurobiología conductual VMTG UNAM**.

No se ha añadido una licencia de código abierto. [NOTICE.md](NOTICE.md) y [docs/PROCEDENCIA.md](docs/PROCEDENCIA.md) documentan las fuentes y los cambios de preparación del repositorio.
