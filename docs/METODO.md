# Entradas y convenciones

## Archivos por stack

La aplicación empareja los archivos mediante identificadores de experimento, animal, corte y región. El patrón de nombre admite, por ejemplo, `HB_04_R1_C5_BLAa_20X` o `SCOP_LiCl_03_R81_C7_BLAp_20X`. Los nombres son ejemplos de formato; no incluyen mediciones.

- Excel de objetos exportado desde Imaris: posiciones XYZ y mediciones de volumen y superficie para el filtro.
- Excel de área/volumen: área en µm², volumen en µm³ y bregma cuando esté disponible.
- ROI de ImageJ: contorno espacial asociado al mismo stack.

Los catálogos de grupos contienen `Experimento`, `Animal`, `Grupo`. El registro de bregma contiene `Experimento`, `Animal y corte`, `Región`, `Bregma`. Los identificadores deben concordar con los nombres de los stacks. Los rangos anatómicos admitidos están configurados en `RANGES`, en `cfos_core.py`.

## Conteo y densidad

El filtro usa explícitamente las columnas de volumen y superficie. Excluye objetos con volumen menor que 50 µm³ **y** razón superficie/volumen mayor que 2.8. Después se aplica la inclusión espacial en la ROI. Las distintas etapas del conteo se conservan para auditoría.

La densidad por animal y región se calcula como:

`densidad_objetos_mm3 = suma(objetos_validos) / suma(volumen_um3) * 1e9`

No debe confundirse un mayor número de objetos con una mayor densidad cuando el volumen muestreado cambia. Los objetos detectados tampoco equivalen automáticamente a células biológicamente validadas.

## Distribución 2D

X e Y se normalizan respecto a la ROI. X aumenta de medial a lateral; Y aumenta de ventral a dorsal. Se usan mapas, centroides y mitades/cuadrantes para describir la distribución. En el análisis posterior se compara la razón de densidades lateral/medial y dorsal/ventral, considerando el volumen de cada mitad.

Una mayor proporción de objetos en una mitad no demuestra una preferencia espacial si esa mitad contiene también más volumen. Los mapas de probabilidad con igual peso por animal y los mapas de densidad responden a preguntas diferentes.

## Distribución 3D

La reconstrucción conserva Z en µm y combina la posición del stack con su bregma. El soporte de los máximos 3D se resume por objetos y animales. En los scripts históricos, `Animales_con_volumen` estima cobertura por solapamiento AP del stack: no verifica por sí solo que cada ROI cubra la posición XY exacta del voxel.

Los máximos de densidad y mapas de diferencia son descriptivos; no son pruebas de significancia voxel a voxel. Las escalas de color de grupo son 0–80 000 objetos/mm³ para 2D y 0–300 000 para 3D. La saturación de color no modifica los valores guardados.

## Inferencia

El animal es la unidad experimental. Stacks, objetos y voxeles son observaciones anidadas. Los scripts en `analysis/` conservan un diseño específico con comparaciones dentro de experimentos, permutaciones estratificadas y familias de corrección Holm. Revisa sus supuestos antes de reutilizarlos con otros grupos o distribuciones de muestra.
