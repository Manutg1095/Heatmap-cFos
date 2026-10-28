# Procedencia del código

Fuentes recuperadas:

- `Heatmap_cFos_V1_2_4_actualizacion_bregma_sin_duplicados.zip`: aplicación, pruebas, dependencias y recurso de inicio.
- `Analisis_cFos_final_reproducible_20260922.zip`: scripts del análisis final y lector XLSX.

Los hashes SHA-256 de los archivos Python originales incluidos se registran en [source_manifest.json](source_manifest.json). Son hashes de las fuentes previas a la preparación del repositorio, no de los archivos adaptados.

Cambios de preparación:

1. Separación de la aplicación (`Heatmap_V1_2/`) y los scripts estadísticos (`analysis/`).
2. Sustitución de los catálogos reales de distribución por plantillas vacías con el mismo esquema. Las pruebas que dependían de la cohorte incrustada usan un catálogo sintético o verifican las plantillas vacías.
3. Sustitución de las rutas absolutas del entorno de trabajo por argumentos `--input`, `--output` y, para figuras, `--results`. Los scripts se ejecutan desde `main()` y pueden importarse sin iniciar un análisis.
4. Corrección de la variable de salida del lanzador zsh: `status` es una variable reservada; se utiliza `heatmap_exit_code`.
5. Inclusión de documentación, dependencias y exclusiones de Git. Se omiten datos, resultados, informes y los constructores específicos de documentos del entorno anterior.

El núcleo científico `cfos_core.py` y la gestión `batch_project.py` se conservan sin modificaciones. Los cálculos estadísticos históricos se preservan; esta preparación no implica una nueva validación estadística del diseño ni una nueva corrida con los animales finales.
