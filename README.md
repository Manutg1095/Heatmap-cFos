# Heatmap cFos

Source code for **Heatmap V1.2, internal version 1.2.4**, for quantifying c-Fos objects and analyzing their spatial distribution in BLAa and BLAp. Includes the desktop application and the final analysis scripts.

The program processes previously segmented objects exported from Imaris, area/volume files, and ImageJ ROIs. It does not segment microscopy images.

## Features

- Experiment and animal registration, animal selection, and stack-level quality control.
- Object counts before filtering, after filtering, and within the ROI.
- Volume-normalized density, with the animal as the primary unit of analysis.
- 2D maps and comparisons of medial–lateral and ventral–dorsal distributions.
- 3D coordinates with Z in µm and anteroposterior position in mm relative to bregma.
- Bregma editing, experiment updates, and consolidation of duplicate Stack_ID entries.
- Export of tables, arrays, and figures for downstream analysis.

## Installation and execution

The code has been tested with Python 3.12. The interface requires Tkinter and a graphical desktop session.

From the repository directory:

```bash
python -m venv .venv
```

Activate the environment using `source .venv/bin/activate` on macOS/Linux or `.venv\Scripts\Activate.ps1` in PowerShell. Then run:

```bash
python -m pip install -r requirements.txt
python Heatmap_V1_2/iniciar.py
```

Run `python -m tkinter` to check your Tkinter installation. On Linux distributions, you may need to install the Tkinter package corresponding to your Python installation. On macOS, the `Heatmap_V1_2/INICIAR_HEATMAP_V1_2.command` launcher is also included for installations with a Conda environment named `cfos`.

## Getting started

1. Open the application and import your catalogs using **Actualizar grupos** (Update groups) and **Actualizar bregmas (catálogo y proyecto)** (Update bregma values: catalog and project). The included templates contain column headers only.
2. Select the object, area/volume, and ROI files for each experiment. Check that the files are correctly matched for each stack.
3. Verify mounting orientation, ROI orientation, bregma, and quality control. Save or update the experiment.
4. Select the animals to include in the analysis and start the run.

By default, the project is saved in `proyecto_cfos/`, alongside the `Heatmap_V1_2/` directory. You can also select an existing project through the interface. Use the `HEATMAP_CFOS_PROJECT` environment variable to specify a different location.

Experimental data, completed catalogs, and results are managed locally. `.gitignore` excludes these paths and data formats, allowing only the two empty templates required at startup. Do not replace these version-controlled templates with actual catalogs; import your catalogs through the interface.

## Anatomical conventions

| Variable | Convention |
| --- | --- |
| Normalized X | 0 = medial; 1 = lateral |
| Normalized Y | 0 = ventral; 1 = dorsal |
| Z | Actual depth in µm |
| AP | Position relative to bregma in mm |
| Primary unit | Animal; sum of valid objects / sum of sampled volumes |

The Y transformation is `Y_DV = 1 - Y_imagen`. Under the mounting scheme used in this study, left-side mounting reverses X, whereas right-side mounting reverses Z; the mounting correction does not introduce an additional Y reversal. Before applying these conventions to another study, verify how the samples were acquired and mounted.

The filter removes an object only when `Volume < 50 µm³` and `Surface / Volume > 2.8` are both satisfied. See [docs/METODO.md](docs/METODO.md) for units, inputs, and limitations of interpretation.

## Statistical analysis code

The [analysis/](analysis/README.md) directory contains the final analysis scripts, with paths configurable through command-line arguments. The scripts retain the study design: BLAa/BLAp and Vehículo (vehicle)/SCOP-HB/LiCl. To run them, also install:

```bash
python -m pip install -r requirements-analysis.txt
```

The accompanying instructions describe the execution order, required files, and assumptions. The repository does not contain study reports or results.

## Tests

```bash
python -m unittest discover -s Heatmap_V1_2 -p "test_*.py" -v
python -m unittest discover -s tests -p "test_*.py" -v
```

The application tests cover filtering, orientation, reconciliation of sector volumes, coordinate export, animal selection, bregma, duplicate entries, and data persistence. The analysis tests verify that modules can be imported without running the study and check the command-line interfaces. See [docs/VALIDACION.md](docs/VALIDACION.md) for the scope of verification.

## Provenance and rights

The original ownership notice is retained: **“Propiedad del Laboratorio de neurobiología conductual VMTG UNAM”** (Property of the Behavioral Neurobiology Laboratory, VMTG UNAM).

No open-source license has been added. [NOTICE.md](NOTICE.md) and [docs/PROCEDENCIA.md](docs/PROCEDENCIA.md) document the sources and the changes made when preparing the repository.
