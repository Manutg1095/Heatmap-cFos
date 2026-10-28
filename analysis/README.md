# Scripts del análisis final

Estos scripts conservan los cálculos del análisis final recuperado. Se adaptaron las rutas y el punto de entrada para ejecutarlos fuera del entorno original. Están configurados para **BLAa**, **BLAp** y los grupos **Vehículo**, **SCOP-HB** y **LiCl**. No constituyen una biblioteca estadística general.

## Entradas

`--input` apunta a una carpeta de corrida exportada por Heatmap que contiene:

- `resultados_cfos_2D_3D.xlsx`, con las hojas originales de QC, animales, sectores, bins 2D y objetos 3D.
- `conteo_objetos.xlsx`, con la hoja `Conteo_por_stack`.
- `LEEME_resultados.txt`.
- `matrices_3D/3D_*.npz`, para evaluar los máximos 3D.

El lector XLSX usa los valores almacenados; no recalcula fórmulas de Excel.

## Ejecución

Desde la raíz del repositorio, sustituye las rutas de ejemplo por tus carpetas:

```bash
python -m pip install -r requirements-analysis.txt
python analysis/cfos_final_analysis.py --input "/ruta/a/corrida" --output "/ruta/a/resultados"
python analysis/cfos_extra_analysis.py --input "/ruta/a/corrida" --output "/ruta/a/resultados"
python analysis/cfos_3d_support.py --input "/ruta/a/corrida" --output "/ruta/a/resultados"
python analysis/make_cfos_figures.py --input "/ruta/a/corrida" --results "/ruta/a/resultados" --output "/ruta/a/figuras"
```

El análisis principal produce los archivos que utiliza el análisis adicional. El script de soporte 3D actualiza `maximos_3D.csv` antes de generar las figuras. Usa una carpeta de salida dedicada: los archivos con el mismo nombre se reemplazan.

| Script | Función |
| --- | --- |
| `cfos_final_analysis.py` | Reconciliación de conteos, densidad por animal, polaridad y centroides 2D, forma espacial, cobertura AP y manifiesto de entradas |
| `cfos_extra_analysis.py` | Sensibilidad a ajustes, advertencias y número de stacks; influencia por animal; mitades y cuadrantes |
| `cfos_3d_support.py` | Soporte de los voxeles de máxima densidad por objetos y animales |
| `make_cfos_figures.py` | Figuras de densidad, polaridad, centroides, mapas 2D y cobertura |
| `read_xlsx_xml.py` | Lectura de valores XLSX directamente desde XML |

## Supuestos y límites heredados

- La unidad de análisis es `Experimento::Animal`, diferenciada por región. Las comparaciones retienen experimentos que contienen ambos grupos comparados.
- Las inferencias principales usan logaritmos de densidad, OLS con errores HC3 y permutación de las etiquetas dentro del experimento. Los intervalos son aproximaciones basadas en t; los valores p de permutación no convierten esos intervalos en intervalos exactos.
- Las pruebas espaciales de mitades/cuadrantes añaden 0.5 a los conteos antes del logaritmo. La densidad global requiere valores positivos; los ceros y volúmenes no positivos necesitan tratamiento explícito antes de aplicar este análisis.
- La malla 2D histórica es de 30 × 30 bins, con varias anchuras de suavizado; el análisis principal de forma usa sigma 1.5. Los mapas con igual peso por animal muestran probabilidad espacial, no densidad absoluta.
- El ajuste AP supone un diseño identificable y observaciones suficientes. La permutación requiere intercambiabilidad dentro de cada experimento; con covariables y grupos pequeños su interpretación debe revisarse para el diseño concreto.
- Se enumeran todas las permutaciones posibles. El tiempo aumenta rápidamente con el tamaño de los grupos; no hay límite automático ni muestreo aleatorio de permutaciones.
- Algunos bucles históricos omiten ajustes que producen excepciones. Revisa el número de comparaciones y permutaciones exportadas; una tabla incompleta no constituye evidencia de ausencia de efecto.
- El script 3D utiliza límites superiores abiertos para identificar objetos en el voxel máximo; los objetos exactamente en el extremo final de un eje pueden diferir del conteo guardado en la matriz. La cobertura de volumen se estima por AP, no por ocupación exacta de cada voxel XY.

Los scripts no incluyen los animales ni los resultados. Las pruebas de preparación verifican su interfaz y portabilidad de importación; no se ha repetido aquí la inferencia con la cohorte final.
