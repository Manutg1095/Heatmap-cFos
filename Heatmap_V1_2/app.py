from __future__ import annotations

import os
import shutil
import time
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from batch_project import (
    COUNTS_NAME,
    DuplicateStackError,
    archive_experiment,
    consolidate_duplicate_stacks,
    experiment_summary,
    generate_roi_orientation_png,
    load_registry,
    refresh_qc,
    refresh_saved_valid_counts,
    resolve_stack_metadata,
    replace_staging_row,
    save_experiment,
    scan_experiment_sets,
    set_animal_analysis_selection,
    set_bregma,
    set_mounting,
    update_saved_bregmas_from_catalog,
    write_counts_workbook,
)
from cfos_core import (
    APP_VERSION,
    META_COLUMNS,
    load_area,
    load_bregma_catalog,
    load_group_catalog,
    numeric_value,
    parse_name,
    run_final,
    selected_for_analysis,
)


ROOT = Path(__file__).resolve().parent
PROJECT_POINTER = Path.home() / "Library" / "Application Support" / "Heatmap_cFos_V1_2" / "project_path.txt"
# V1.2 editable: ningún experimento guardado queda congelado.
FROZEN_EXPERIMENTS = set()
MASTER_PROJECT_MARKER = "PROYECTO_MAESTRO_V1_2.txt"


def contains_saved_project(path):
    path = Path(path).expanduser()
    return (
        (path / "registro_general_stacks.xlsx").is_file()
        or (path / "metadata_stacks.xlsx").is_file()
        or (path / "experimentos").is_dir()
    )


def stable_project_path(root=ROOT):
    return Path(root).resolve().parent / "proyecto_cfos"


def is_master_project(path):
    path = Path(path).expanduser()
    return contains_saved_project(path) and (path / MASTER_PROJECT_MARKER).is_file()


def is_project_inside_app(path):
    """Reconoce proyectos ligados a una carpeta versionada de la aplicación."""
    path = Path(path).expanduser().resolve()
    return path.name == "proyecto_cfos" and path.parent.name.casefold().startswith("heatmap_v")


def copy_project_to_stable(source, root=ROOT):
    """Copia sin borrar el proyecto legado a la ubicación estable recomendada."""
    source = Path(source).expanduser().resolve()
    destination = stable_project_path(root)
    if source == destination:
        return destination
    if destination.exists() and any(destination.iterdir()):
        if contains_saved_project(destination):
            return destination.resolve()
        raise FileExistsError(
            f"La ubicación estable no está vacía: {destination}. "
            "Muévela o selecciona otra carpeta antes de migrar."
        )
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, destination, dirs_exist_ok=True)
    return destination.resolve()


def resolve_project_path(root=ROOT, pointer=PROJECT_POINTER):
    """Localiza datos persistentes sin ligarlos a la carpeta de la app."""
    root = Path(root).resolve()
    override = os.environ.get("HEATMAP_CFOS_PROJECT", "").strip()
    if override:
        return Path(override).expanduser().resolve()
    stable = stable_project_path(root)
    # La distribución V1.2 incluye un proyecto maestro propio. Tiene prioridad
    # sobre cualquier puntero guardado por versiones preliminares, de modo que
    # «Ejecutar» nunca vuelva silenciosamente a una app o carpeta anterior.
    if is_master_project(stable):
        return stable.resolve()
    try:
        if Path(pointer).is_file():
            saved = Path(Path(pointer).read_text(encoding="utf-8").strip()).expanduser()
            if saved.exists():
                if is_project_inside_app(saved) and contains_saved_project(stable):
                    return stable.resolve()
                return saved.resolve()
    except OSError:
        pass

    candidates = [root / "proyecto_cfos", stable]
    try:
        candidates.extend(
            child / "proyecto_cfos"
            for child in root.parent.iterdir()
            if child.is_dir() and child != root
        )
    except OSError:
        pass
    unique = []
    for candidate in candidates:
        candidate = candidate.resolve()
        if candidate not in unique and contains_saved_project(candidate):
            unique.append(candidate)
    if len(unique) == 1:
        return unique[0]
    if contains_saved_project(root / "proyecto_cfos"):
        return (root / "proyecto_cfos").resolve()
    if contains_saved_project(stable):
        return stable.resolve()
    return stable.resolve()


def configure_project(path, persist=False):
    global PROJECT, PROJECT_CATALOGS, GROUP_CATALOG, BREGMA_CATALOG
    PROJECT = Path(path).expanduser().resolve()
    PROJECT_CATALOGS = PROJECT / "catalogos"
    GROUP_CATALOG = PROJECT_CATALOGS / "asignacion_grupos.xlsx"
    BREGMA_CATALOG = PROJECT_CATALOGS / "registro_bregma.xlsx"
    if persist:
        try:
            PROJECT_POINTER.parent.mkdir(parents=True, exist_ok=True)
            PROJECT_POINTER.write_text(str(PROJECT), encoding="utf-8")
        except OSError:
            pass
    return PROJECT


configure_project(resolve_project_path())
BUNDLED_GROUPS = ROOT / "catalogos" / "asignacion_grupos.xlsx"
BUNDLED_BREGMA = ROOT / "catalogos" / "registro_bregma.xlsx"
SPLASH_IMAGE = ROOT / "assets" / "splash_v11.png"
SPLASH_MIN_MS = 7000
PRODUCT_NAME = "Heatmap V1.2"
OWNERSHIP_TEXT = "Propiedad del Laboratorio de neurobiología conductual VMTG UNAM"


STAGING_COLUMNS = [
    "Stack_ID", "Grupo", "Bregma_mm", "Region", "Objetos_XYZ",
    "Objetos_incluidos_filtro", "Objetos_validos", "Area_um2", "Volumen_um3",
    "Montado_lado_derecho", "Revision_ROI",
    "QC", "Usar_mapa_grupal", "Detalle_QC",
]
STAGING_HEADINGS = {
    "Montado_lado_derecho": "Montado de lado derecho",
}
PROJECT_COLUMNS = [
    "Incluir_analisis", "Stack_ID", "Experimento", "Animal", "Grupo", "Region", "Bregma_mm",
    "Montado_lado_derecho", "Objetos_validos", "QC", "Usar_mapa_grupal", "Actualizado",
]
PROJECT_HEADINGS = {"Incluir_analisis": "Analizar"}


def format_number(value):
    number = numeric_value(value)
    return "" if number is None else f"{number:g}"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.title(f"{PRODUCT_NAME} — Versión {APP_VERSION}")
        self.geometry("1420x920")
        self.minsize(1120, 760)
        self._splash_started = time.monotonic()
        self._show_splash()
        self.after(60, self._initialize)

    def _show_splash(self):
        self.splash = tk.Toplevel(self)
        self.splash.overrideredirect(True)
        self.splash.configure(bg="#061020")
        width, height = 680, 790
        x = max(0, (self.winfo_screenwidth() - width) // 2)
        y = max(0, (self.winfo_screenheight() - height) // 2)
        self.splash.geometry(f"{width}x{height}+{x}+{y}")
        tk.Label(
            self.splash, text=PRODUCT_NAME, bg="#061020", fg="#F2F7FF",
            font=("Helvetica Neue", 25, "bold"),
        ).pack(pady=(13, 7))
        if SPLASH_IMAGE.exists():
            try:
                self._splash_original = tk.PhotoImage(file=str(SPLASH_IMAGE))
                self.splash_photo = self._splash_original.subsample(2, 2)
                tk.Label(self.splash, image=self.splash_photo, borderwidth=0, bg="#061020").pack()
            except tk.TclError:
                self._splash_fallback()
        else:
            self._splash_fallback()
        self.splash_status = tk.StringVar(value="Preparando el proyecto por experimentos…")
        tk.Label(
            self.splash, textvariable=self.splash_status, bg="#061020", fg="#91E9DE",
            font=("Helvetica Neue", 10),
        ).pack(pady=(7, 2))
        progress = ttk.Progressbar(self.splash, mode="indeterminate", length=300)
        progress.pack(pady=(0, 6))
        progress.start(10)
        tk.Label(
            self.splash, text=OWNERSHIP_TEXT, bg="#061020", fg="#D4DDED",
            font=("Helvetica Neue", 10),
        ).pack(side="bottom", pady=(2, 11))
        self.splash.lift()

    def _splash_fallback(self):
        tk.Label(
            self.splash, text="c-Fos · BLA · análisis espacial 2D/3D",
            bg="#061020", fg="#62E6D7", font=("Helvetica Neue", 18), height=20,
        ).pack(fill="both", expand=True)

    def _initialize(self):
        try:
            self._bootstrap_catalogs()
            self.groups = load_group_catalog(GROUP_CATALOG)
            self.bregmas = load_bregma_catalog(BREGMA_CATALOG)
            bregma_sync = update_saved_bregmas_from_catalog(PROJECT, self.bregmas)
            self.data = bregma_sync["Data"]
            self.staging = pd.DataFrame(columns=META_COLUMNS)
            self.folder_vars = {
                key: tk.StringVar()
                for key in ("Archivo_objetos", "Archivo_area", "Archivo_ROI")
            }
            self.catalog_status = tk.StringVar()
            self.scan_status = tk.StringVar(value="Selecciona las tres carpetas de un experimento.")
            self.project_status = tk.StringVar()
            self.detected_experiment = tk.StringVar(value="—")
            self.edit_vars = {
                "Bregma_mm": tk.StringVar(value=""),
                "Montado_lado_derecho": tk.StringVar(value=""),
                "Revision_ROI": tk.StringVar(value="Pendiente"),
                "QC": tk.StringVar(value="Revisar"),
                "Usar_mapa_grupal": tk.StringVar(value="Si"),
                "Notas": tk.StringVar(),
            }
            self._build_interface()
            self._refresh_catalog_status()
            self._refresh_staging()
            self._refresh_project()
            self.splash_status.set(
                f"Listo: {len(self.groups)} animales y {len(self.bregmas)} cortes con bregma · "
                f"{bregma_sync['Bregmas_cambiados']} valor(es) sincronizado(s)"
            )
            self._finish_splash_after_minimum()
        except Exception as exc:
            self.splash.destroy()
            self.deiconify()
            messagebox.showerror("No se pudo iniciar Heatmap V1.2", str(exc))

    def _finish_splash_after_minimum(self):
        elapsed_ms = int((time.monotonic() - self._splash_started) * 1000)
        self.after(max(0, SPLASH_MIN_MS - elapsed_ms), self._reveal_main_window)

    def _reveal_main_window(self):
        self.splash.destroy()
        self.deiconify()
        self.lift()
        if is_project_inside_app(PROJECT):
            self.after(150, self._offer_project_migration)

    def _offer_project_migration(self):
        destination = stable_project_path(ROOT)
        answer = messagebox.askyesno(
            "Proyecto dentro de una versión anterior",
            f"El proyecto activo está dentro de una carpeta de la aplicación:\n{PROJECT}\n\n"
            "Si actualizas o eliminas esa versión, las rutas pueden romperse y las corridas "
            "seguirán guardándose en la app anterior.\n\n"
            f"¿Copiar ahora el proyecto completo a la ubicación estable?\n{destination}",
        )
        if not answer:
            return
        try:
            candidate = copy_project_to_stable(PROJECT, ROOT)
            self._activate_project(candidate)
            messagebox.showinfo(
                "Proyecto migrado",
                f"La copia estable quedó activa en:\n{candidate}\n\n"
                "El proyecto anterior no se eliminó. Las próximas corridas se guardarán aquí.",
            )
        except Exception as exc:
            messagebox.showerror("No se pudo migrar el proyecto", str(exc))

    def _activate_project(self, candidate):
        configure_project(candidate, persist=True)
        self._bootstrap_catalogs()
        self.groups = load_group_catalog(GROUP_CATALOG)
        self.bregmas = load_bregma_catalog(BREGMA_CATALOG)
        self.data = update_saved_bregmas_from_catalog(PROJECT, self.bregmas)["Data"]
        self._refresh_catalog_status()
        self._refresh_project()

    def _bootstrap_catalogs(self):
        PROJECT_CATALOGS.mkdir(parents=True, exist_ok=True)
        if not GROUP_CATALOG.exists():
            shutil.copy2(BUNDLED_GROUPS, GROUP_CATALOG)
        if not BREGMA_CATALOG.exists():
            shutil.copy2(BUNDLED_BREGMA, BREGMA_CATALOG)

    def _build_interface(self):
        header = tk.Frame(self, bg="#061020", padx=14, pady=10)
        header.pack(fill="x")
        tk.Label(
            header, text=PRODUCT_NAME, bg="#061020", fg="white",
            font=("Helvetica Neue", 21, "bold"),
        ).pack(side="left")
        tk.Label(
            header, text=f"Versión interna {APP_VERSION} · carga por experimentos",
            bg="#061020", fg="#70E1D2", font=("Helvetica Neue", 10),
        ).pack(side="left", padx=14)
        tk.Label(
            header, textvariable=self.catalog_status, bg="#061020", fg="#D3DEF0",
            font=("Helvetica Neue", 10),
        ).pack(side="right")

        catalogs = ttk.Frame(self, padding=(10, 7))
        catalogs.pack(fill="x")
        ttk.Button(
            catalogs, text="Actualizar grupos", command=lambda: self.import_catalog("groups")
        ).pack(side="left", padx=(0, 7))
        ttk.Button(
            catalogs, text="Actualizar bregmas (catálogo y proyecto)",
            command=lambda: self.import_catalog("bregma")
        ).pack(side="left")
        ttk.Label(
            catalogs,
            text=(
                "El catálogo actualiza los Bregmas guardados; una edición manual conserva prioridad."
            ),
        ).pack(side="left", padx=14)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        self.load_tab = ttk.Frame(self.notebook)
        self.project_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.load_tab, text="1. Cargar y revisar experimento")
        self.notebook.add(self.project_tab, text="2. Proyecto guardado y análisis final")
        self._build_load_tab()
        self._build_project_tab()

        tk.Label(
            self, text=OWNERSHIP_TEXT, bg="#061020", fg="#D4DDED",
            font=("Helvetica Neue", 9), pady=5,
        ).pack(fill="x", side="bottom")

    def _build_load_tab(self):
        folders = ttk.LabelFrame(
            self.load_tab, text="A. Selecciona el set completo del experimento", padding=9
        )
        folders.pack(fill="x", padx=7, pady=(8, 4))
        rows = [
            ("Archivo_objetos", "Carpeta de objetos Imaris"),
            ("Archivo_area", "Carpeta de área/volumen"),
            ("Archivo_ROI", "Carpeta de ROIs ImageJ"),
        ]
        for row, (key, label) in enumerate(rows):
            ttk.Label(folders, text=label, width=27).grid(row=row, column=0, sticky="w", pady=2)
            ttk.Entry(folders, textvariable=self.folder_vars[key], state="readonly").grid(
                row=row, column=1, sticky="ew", pady=2
            )
            ttk.Button(folders, text="Seleccionar carpeta", command=lambda key=key: self.pick_folder(key)).grid(
                row=row, column=2, padx=(7, 0), pady=2
            )
        folders.columnconfigure(1, weight=1)
        ttk.Button(folders, text="Escanear y generar pre-registro", command=self.scan).grid(
            row=0, column=3, rowspan=2, padx=(14, 0), sticky="nsew"
        )
        ttk.Button(folders, text="Limpiar pre-registro", command=self.clear_staging).grid(
            row=2, column=3, padx=(14, 0), sticky="ew"
        )
        ttk.Button(
            folders,
            text="Abrir PNG inicial de ROIs",
            command=self.open_initial_roi_png,
        ).grid(row=1, column=4, padx=(7, 0), sticky="ew")
        ttk.Button(
            folders,
            text="Actualizar / abrir PNG corregido",
            command=self.open_roi_orientation_png,
        ).grid(row=2, column=4, padx=(7, 0), sticky="ew")

        defaults = ttk.LabelFrame(self.load_tab, text="B. Clasificar el montaje después de ver el PNG inicial", padding=9)
        defaults.pack(fill="x", padx=7, pady=4)
        ttk.Label(defaults, text="Experimento detectado:").pack(side="left")
        ttk.Label(defaults, textvariable=self.detected_experiment, foreground="#0B6E69").pack(
            side="left", padx=(5, 18)
        )
        ttk.Label(
            defaults,
            text=(
                "Todos son hemisferio izquierdo. No=montado del lado izquierdo → corrige X; "
                "Sí=montado de lado derecho → corrige Z. Y no cambia."
            ),
            wraplength=850,
        ).pack(side="left", padx=(0, 18))
        ttk.Button(
            defaults,
            text="Marcar todos los ROIs como revisados",
            command=self.mark_all_rois_reviewed,
        ).pack(side="right", padx=(10, 0))

        ttk.Label(self.load_tab, textvariable=self.scan_status).pack(
            fill="x", padx=12, pady=(1, 4)
        )
        self._build_staging_table()
        self._build_selected_editor()

        actions = ttk.Frame(self.load_tab, padding=(7, 7))
        actions.pack(fill="x")
        ttk.Button(
            actions, text="Guardar / actualizar experimento", command=self.save_current_experiment
        ).pack(side="left")
        ttk.Label(
            actions,
            text=(
                "Guardar actualiza el registro general y conteo_objetos.xlsx; "
                "los heatmaps se generan solamente en el análisis final."
            ),
        ).pack(side="left", padx=14)

    def _build_staging_table(self):
        frame = ttk.LabelFrame(self.load_tab, text="C. Pre-registro para revisión", padding=5)
        frame.pack(fill="both", expand=True, padx=7, pady=4)
        self.staging_tree = ttk.Treeview(frame, columns=STAGING_COLUMNS, show="headings", height=13)
        widths = {
            "Stack_ID": 220, "Detalle_QC": 330, "Grupo": 100, "Bregma_mm": 85,
            "Volumen_um3": 115, "Area_um2": 105, "Montado_lado_derecho": 170,
        }
        for column in STAGING_COLUMNS:
            self.staging_tree.heading(column, text=STAGING_HEADINGS.get(column, column))
            self.staging_tree.column(column, width=widths.get(column, 105), anchor="w")
        # El pre-registro usa una presentación neutra: fondo blanco y texto negro.
        for tag in ("ok", "review", "exclude"):
            self.staging_tree.tag_configure(tag, background="#FFFFFF", foreground="#000000")
        self.staging_tree.bind("<<TreeviewSelect>>", self.load_selected_row)
        self.staging_tree.bind("<Double-1>", self.toggle_mounting_cell)
        vertical = ttk.Scrollbar(frame, orient="vertical", command=self.staging_tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=self.staging_tree.xview)
        self.staging_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.staging_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _build_selected_editor(self):
        frame = ttk.LabelFrame(self.load_tab, text="D. Editar fila seleccionada", padding=8)
        frame.pack(fill="x", padx=7, pady=4)
        ttk.Label(frame, text="Bregma (mm)").pack(side="left", padx=(0, 4))
        ttk.Entry(frame, textvariable=self.edit_vars["Bregma_mm"], width=9).pack(
            side="left", padx=(0, 10)
        )
        fields = [
            ("Montado_lado_derecho", "Montado de lado derecho", ["", "No", "Si"], 8),
            ("QC", "QC", ["OK", "Revisar", "Excluir"], 10),
            ("Usar_mapa_grupal", "Mapa grupal", ["Si", "No"], 8),
            ("Revision_ROI", "Revisión ROI", ["Pendiente", "Revisado"], 10),
        ]
        for key, label, values, width in fields:
            ttk.Label(frame, text=label).pack(side="left", padx=(0, 4))
            ttk.Combobox(
                frame, textvariable=self.edit_vars[key], values=values,
                state="readonly", width=width,
            ).pack(side="left", padx=(0, 10))
        ttk.Label(frame, text="Notas").pack(side="left", padx=(10, 4))
        ttk.Entry(frame, textvariable=self.edit_vars["Notas"], width=28).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(frame, text="Aplicar a fila", command=self.apply_selected).pack(
            side="right", padx=(8, 0)
        )

    def _build_project_tab(self):
        actions = ttk.LabelFrame(self.project_tab, text="Proyecto completo", padding=9)
        actions.pack(fill="x", padx=7, pady=(8, 4))
        ttk.Label(actions, textvariable=self.project_status).pack(side="left")
        ttk.Button(actions, text="Actualizar vista", command=self._refresh_project).pack(
            side="right", padx=(7, 0)
        )
        ttk.Button(actions, text="Ejecutar análisis completo", command=self.run).pack(side="right")
        ttk.Button(
            actions, text="Actualizar sólo objetos válidos",
            command=self.refresh_valid_counts_only,
        ).pack(side="right", padx=(7, 0))
        ttk.Button(
            actions,
            text="Recuperar / cambiar proyecto",
            command=self.select_existing_project,
        ).pack(side="right", padx=(7, 0))

        summary_frame = ttk.LabelFrame(
            self.project_tab, text="Experimentos guardados", padding=5
        )
        summary_frame.pack(fill="x", padx=7, pady=4)
        summary_columns = [
            "Experimento", "N_stacks", "N_animales", "BLAa", "BLAp", "QC_OK", "Revisar", "Excluir"
        ]
        self.experiment_tree = ttk.Treeview(
            summary_frame, columns=summary_columns, show="headings", height=5
        )
        for column in summary_columns:
            self.experiment_tree.heading(column, text=column)
            self.experiment_tree.column(column, width=125, anchor="center")
        self.experiment_tree.bind("<Double-1>", lambda _event: self.load_saved_experiment())
        self.experiment_tree.pack(fill="x", side="left", expand=True)
        experiment_actions = ttk.Frame(summary_frame)
        experiment_actions.pack(side="right", fill="y", padx=8)
        ttk.Button(
            experiment_actions, text="Editar experimento\nseleccionado",
            command=self.load_saved_experiment,
        ).pack(fill="x", pady=(0, 6))
        ttk.Button(
            experiment_actions, text="Eliminar experimento\nseleccionado",
            command=self.delete_saved_experiment,
        ).pack(fill="x")

        frame = ttk.LabelFrame(self.project_tab, text="Registro general de stacks", padding=5)
        frame.pack(fill="both", expand=True, padx=7, pady=4)
        self.project_tree = ttk.Treeview(frame, columns=PROJECT_COLUMNS, show="headings")
        for column in PROJECT_COLUMNS:
            self.project_tree.heading(column, text=PROJECT_HEADINGS.get(column, column))
            self.project_tree.column(
                column,
                width=68 if column == "Incluir_analisis" else 170 if column in {"Stack_ID", "Montado_lado_derecho"} else 115,
                anchor="center" if column == "Incluir_analisis" else "w",
            )
        self.project_tree.bind("<Double-1>", self.toggle_analysis_cell)
        self.project_tree.bind("<space>", self.toggle_selected_animal)
        vertical = ttk.Scrollbar(frame, orient="vertical", command=self.project_tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=self.project_tree.xview)
        self.project_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.project_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        ttk.Label(
            self.project_tab,
            text=(
                "Doble clic en Analizar (o barra espaciadora) para incluir/excluir al animal completo. "
                "La selección queda guardada; el análisis final crea una corrida independiente."
            ),
        ).pack(fill="x", padx=10, pady=(2, 8))

    def _refresh_catalog_status(self):
        self.catalog_status.set(
            f"Catálogos: {len(self.groups)} animales · {len(self.bregmas)} cortes"
        )

    def select_existing_project(self):
        selected = filedialog.askdirectory(
            title="Selecciona la carpeta proyecto_cfos que contiene tus datos",
            initialdir=str(PROJECT.parent),
        )
        if not selected:
            return
        candidate = Path(selected).expanduser()
        nested = candidate / "proyecto_cfos"
        if not contains_saved_project(candidate) and contains_saved_project(nested):
            candidate = nested
        if not contains_saved_project(candidate):
            return messagebox.showerror(
                "Proyecto no reconocido",
                "La carpeta seleccionada no contiene registro_general_stacks.xlsx, "
                "metadata_stacks.xlsx ni la carpeta experimentos.\n\n"
                "Selecciona directamente tu carpeta proyecto_cfos anterior.",
            )
        master = stable_project_path(ROOT)
        if is_master_project(master) and candidate.resolve() != master.resolve():
            return messagebox.showinfo(
                "Proyecto maestro V1.2 activo",
                "Esta distribución ya contiene una copia propia y congelada de HB_02, "
                "SCOP_LiCl_02 y SCOP_LiCl_03. Para evitar que «Ejecutar» vuelva a escribir "
                "en una app anterior, el proyecto activo no se cambiará.\n\n"
                f"Proyecto activo:\n{master}",
            )
        try:
            if is_project_inside_app(candidate):
                destination = stable_project_path(ROOT)
                migrate = messagebox.askyesno(
                    "Proyecto dentro de la aplicación",
                    f"Ese proyecto está ligado a una versión anterior.\n\n"
                    f"¿Copiarlo y usar la ubicación estable?\n{destination}",
                )
                if not migrate:
                    return
                candidate = copy_project_to_stable(candidate, ROOT)
            self._activate_project(candidate)
            messagebox.showinfo(
                "Proyecto recuperado",
                f"Se cargaron {len(self.data)} stacks desde:\n{PROJECT}\n\n"
                "Las corridas anteriores no se modificaron. Ejecuta nuevamente el análisis "
                "completo para crear una salida nueva con Y=0 ventral y Y=1 dorsal.",
            )
        except Exception as exc:
            messagebox.showerror("No se pudo recuperar el proyecto", str(exc))

    def import_catalog(self, kind):
        path = filedialog.askopenfilename(filetypes=[("CSV o Excel", "*.csv *.xlsx *.xls")])
        if not path:
            return
        try:
            if kind == "groups":
                frame = load_group_catalog(path)
                target = GROUP_CATALOG
                label = "asignaciones de grupo"
            else:
                frame = load_bregma_catalog(path)
                target = BREGMA_CATALOG
                label = "cortes con bregma"
            if Path(path).resolve() != target.resolve():
                shutil.copy2(path, target)
            if kind == "groups":
                self.groups = frame
                sync = None
            else:
                self.bregmas = frame
                sync = update_saved_bregmas_from_catalog(PROJECT, self.bregmas)
                self.data = sync["Data"]
            self._refresh_catalog_status()
            self._reapply_catalogs_to_staging(kind)
            self._refresh_project()
            detail = f"Se validaron {len(frame)} {label}."
            if sync is not None:
                detail += (
                    f"\n\nBregmas cambiados en el proyecto: {sync['Bregmas_cambiados']}"
                    f"\nCoincidencias: {sync['Coincidencias']}"
                    f"\nEdiciones manuales conservadas: {sync['Omitidos_manual']}"
                    f"\nStacks sin coincidencia: {sync['Sin_coincidencia']}"
                    f"\nFilas duplicadas reparadas: {sync['Duplicados_reparados']}"
                )
                if sync["Stack_ID_consolidados"]:
                    detail += "\nReparados: " + ", ".join(sync["Stack_ID_consolidados"][:8])
            messagebox.showinfo("Catálogo actualizado", detail)
        except Exception as exc:
            messagebox.showerror("No se pudo actualizar el catálogo", str(exc))

    def _reapply_catalogs_to_staging(self, kind="all"):
        if self.staging.empty:
            return
        self.staging, _ = consolidate_duplicate_stacks(PROJECT, self.staging)
        rows = []
        for _, row in self.staging.iterrows():
            record = row.to_dict()
            parsed = parse_name(record["Stack_ID"])
            area_info = {"Bregma_area_mm": record.get("Bregma_area_mm", "")}
            resolved = resolve_stack_metadata(
                parsed,
                self.groups,
                self.bregmas,
                area_info,
                prefer_catalog=kind in {"bregma", "all"},
            )
            if kind in {"groups", "all"}:
                record["Grupo"] = resolved["Grupo"]
                record["Fuente_grupo"] = resolved["Fuente_grupo"]
            if (
                kind in {"bregma", "all"}
                and str(record.get("Fuente_bregma", "")).strip() != "Edición manual"
                and resolved["Fuente_bregma"] == "Registro maestro actualizado"
            ):
                record = set_bregma(
                    record,
                    resolved["Bregma_mm"],
                    source="Registro maestro actualizado",
                )
            rows.append(refresh_qc(record))
        self.staging = pd.DataFrame(rows).reindex(columns=META_COLUMNS, fill_value="")
        self._refresh_staging()

    def pick_folder(self, key):
        path = filedialog.askdirectory()
        if path:
            self.folder_vars[key].set(path)

    def scan(self):
        if not all(variable.get() for variable in self.folder_vars.values()):
            return messagebox.showerror("Faltan carpetas", "Selecciona objetos, área/volumen y ROIs.")
        self.scan_status.set("Escaneando, emparejando y contando objetos…")
        self.update_idletasks()
        try:
            staging, report = scan_experiment_sets(
                self.folder_vars["Archivo_objetos"].get(),
                self.folder_vars["Archivo_area"].get(),
                self.folder_vars["Archivo_ROI"].get(),
                self.groups,
                self.bregmas,
            )
            self.staging = staging
            self.detected_experiment.set(report["Experimento"])
            ok = int(self.staging["QC"].eq("OK").sum())
            review = int(self.staging["QC"].eq("Revisar").sum())
            ignored = report["Ignorados"]
            self.scan_status.set(
                f"{report['Experimento']}: {len(self.staging)} stacks · {ok} OK · "
                f"{review} por revisar · {ignored} archivo(s) no reconocidos ignorados."
            )
            self.open_initial_roi_png()
        except Exception as exc:
            self.scan_status.set(str(exc))
            messagebox.showerror("No se pudo generar el pre-registro", str(exc))

    def clear_staging(self):
        self.staging = pd.DataFrame(columns=META_COLUMNS)
        self.detected_experiment.set("—")
        self.scan_status.set("Selecciona las tres carpetas de un experimento.")
        self._refresh_staging()

    def _create_roi_orientation_png(self, open_file=False, show_message=True, apply_corrections=True):
        if self.staging.empty:
            if show_message:
                messagebox.showinfo("Sin pre-registro", "Primero escanea o carga un experimento.")
            return None
        try:
            output, report = generate_roi_orientation_png(
                PROJECT,
                self.staging,
                self.detected_experiment.get() if self.detected_experiment.get() != "—" else None,
                apply_corrections=apply_corrections,
            )
            errors = report["Errores"]
            self.scan_status.set(
                f"PNG de orientación: {report['ROIs_visibles']}/{report['Total']} ROIs visibles · {output}"
            )
            if open_file:
                webbrowser.open(output.resolve().as_uri())
            if errors and show_message:
                messagebox.showwarning(
                    "PNG generado con advertencias",
                    f"Se mostraron {report['ROIs_visibles']} de {report['Total']} ROIs.\n\n"
                    + "\n".join(errors[:8]),
                )
            return output
        except Exception as exc:
            if show_message:
                messagebox.showerror("No se pudo crear el PNG de ROIs", str(exc))
            return None

    def open_roi_orientation_png(self):
        return self._create_roi_orientation_png(open_file=True, show_message=True, apply_corrections=True)

    def open_initial_roi_png(self):
        return self._create_roi_orientation_png(open_file=True, show_message=True, apply_corrections=False)

    def mark_all_rois_reviewed(self):
        if self.staging.empty:
            return messagebox.showinfo("Sin pre-registro", "Primero escanea o carga un experimento.")
        if self._staging_is_frozen():
            return self._show_frozen_notice()
        answer = messagebox.askyesno(
            "Confirmar revisión de ROIs",
            "Confirma que revisaste el PNG y que la orientación de todos los ROIs es correcta.\n\n"
            "¿Marcar todos como Revisado?",
        )
        if not answer:
            return
        rows = []
        for _, row in self.staging.iterrows():
            record = row.to_dict()
            record["Revision_ROI"] = "Revisado"
            rows.append(refresh_qc(record))
        self.staging = pd.DataFrame(rows).reindex(columns=META_COLUMNS, fill_value="")
        self._refresh_staging()
        self._create_roi_orientation_png(open_file=False, show_message=False, apply_corrections=True)

    def _refresh_staging(self):
        self.staging = self.staging.reindex(
            columns=META_COLUMNS, fill_value=""
        ).reset_index(drop=True)
        if not hasattr(self, "staging_tree"):
            return
        self.staging_tree.delete(*self.staging_tree.get_children())
        for index, row in self.staging.reset_index(drop=True).iterrows():
            qc = str(row.get("QC", "")).casefold()
            tag = "ok" if qc == "ok" else "exclude" if qc == "excluir" else "review"
            values = []
            for column in STAGING_COLUMNS:
                value = row.get(column, "")
                if column in {"Bregma_mm", "Area_um2", "Volumen_um3"}:
                    value = format_number(value)
                elif column == "Montado_lado_derecho":
                    value = str(value).strip() or "Sin clasificar"
                values.append(value)
            self.staging_tree.insert("", "end", iid=str(index), values=values, tags=(tag,))

    def load_selected_row(self, _event=None):
        selected = self.staging_tree.selection()
        if not selected:
            return
        index = int(selected[0])
        row = self.staging.iloc[index]
        for key, variable in self.edit_vars.items():
            value = row.get(key, "")
            variable.set(format_number(value) if key == "Bregma_mm" else str(value))

    def toggle_mounting_cell(self, event):
        if self._staging_is_frozen():
            return self._show_frozen_notice()
        row_id = self.staging_tree.identify_row(event.y)
        column_id = self.staging_tree.identify_column(event.x)
        if not row_id or not column_id:
            return
        column_index = int(column_id.removeprefix("#")) - 1
        if column_index < 0 or STAGING_COLUMNS[column_index] != "Montado_lado_derecho":
            return
        index = int(row_id)
        current = str(self.staging.iloc[index].get("Montado_lado_derecho", "No")).casefold()
        record = set_mounting(self.staging.iloc[index].to_dict(), "No" if current in {"si", "sí"} else "Si")
        self.staging = replace_staging_row(self.staging, index, record)
        self._refresh_staging()
        self.staging_tree.selection_set(str(index))
        self.load_selected_row()
        self._create_roi_orientation_png(open_file=False, show_message=False, apply_corrections=True)

    def apply_selected(self):
        if self._staging_is_frozen():
            return self._show_frozen_notice()
        selected = self.staging_tree.selection()
        if not selected:
            return messagebox.showinfo("Sin fila", "Selecciona una fila del pre-registro.")
        index = int(selected[0])
        record = self.staging.iloc[index].to_dict()
        requested_qc = self.edit_vars["QC"].get()
        for key, variable in self.edit_vars.items():
            if key in {"Bregma_mm", "Montado_lado_derecho"}:
                continue
            record[key] = variable.get()
        try:
            requested_bregma = numeric_value(self.edit_vars["Bregma_mm"].get())
            previous_bregma = numeric_value(record.get("Bregma_mm"))
            if requested_bregma is None:
                raise ValueError("Escribe un Bregma numérico en milímetros.")
            if previous_bregma is None or abs(requested_bregma - previous_bregma) > 1e-9:
                record = set_bregma(record, requested_bregma)
        except ValueError as exc:
            return messagebox.showerror("Bregma no válido", str(exc))
        record = set_mounting(record, self.edit_vars["Montado_lado_derecho"].get())
        if requested_qc != "Excluir":
            record["QC"] = ""
        record = refresh_qc(record)
        self.staging = replace_staging_row(self.staging, index, record)
        self._refresh_staging()
        self.staging_tree.selection_set(str(index))
        self._create_roi_orientation_png(open_file=False, show_message=False)

    def save_current_experiment(self):
        if self.staging.empty:
            return messagebox.showerror("Sin pre-registro", "Primero escanea o carga un experimento.")
        if self._staging_is_frozen():
            return self._show_frozen_notice()
        self._create_roi_orientation_png(open_file=False, show_message=False)
        review = self.staging[self.staging["QC"].eq("Revisar")]
        if not review.empty:
            answer = messagebox.askyesno(
                "Hay filas por revisar",
                f"Hay {len(review)} stack(s) con QC=Revisar. Puedes guardarlos para continuar después, "
                "pero el análisis final quedará bloqueado hasta corregirlos o marcarlos como Excluir.\n\n"
                "¿Quieres guardar el experimento de todos modos?",
            )
            if not answer:
                return
        try:
            result = save_experiment(PROJECT, self.staging, overwrite=False)
        except DuplicateStackError as exc:
            answer = messagebox.askyesno(
                "Stacks ya registrados",
                f"{len(exc.stack_ids)} Stack_ID ya existen. ¿Quieres reemplazar sus registros?\n\n"
                "Los archivos anteriores permanecerán en sus carpetas versionadas.",
            )
            if not answer:
                return
            try:
                result = save_experiment(PROJECT, self.staging, overwrite=True)
            except Exception as second_exc:
                return messagebox.showerror("No se pudo guardar el experimento", str(second_exc))
        except Exception as exc:
            return messagebox.showerror("No se pudo guardar el experimento", str(exc))

        self.data = result["Data"]
        self.staging = self.data[
            self.data["Experimento"].eq(result["Experimento"])
        ].copy().reset_index(drop=True)
        self._refresh_staging()
        self._refresh_project()
        messagebox.showinfo(
            "Experimento guardado",
            f"{result['Experimento']}: {result['N_stacks']} stacks registrados.\n\n"
            f"Se actualizó {COUNTS_NAME}. Los heatmaps todavía no se ejecutaron.",
        )

    def refresh_valid_counts_only(self):
        if self.data.empty:
            return messagebox.showerror("Sin datos", "No hay stacks guardados para actualizar.")
        answer = messagebox.askyesno(
            "Actualizar sólo objetos válidos",
            "Se recalculará únicamente Objetos_validos usando los Excel de objetos y las ROI ya guardadas.\n\n"
            "No se modificarán bregma, grupo, área, volumen, orientación, QC, archivos ni otros metadatos.\n\n"
            "¿Continuar?",
        )
        if not answer:
            return
        try:
            result = refresh_saved_valid_counts(PROJECT)
            self._refresh_project()
            audit = result["Auditoria"]
            changed = audit[audit["Diferencia"].fillna(0).ne(0)] if not audit.empty else audit
            examples = []
            for _, row in changed.head(6).iterrows():
                before = row["Objetos_validos_antes"]
                before_text = "—" if pd.isna(before) else str(int(before))
                examples.append(f"{row['Stack_ID']}: {before_text} → {int(row['Objetos_validos_nuevo'])}")
            detail = "\n".join(examples)
            if len(changed) > 6:
                detail += f"\n… y {len(changed) - 6} cambio(s) más."
            errors = result["Errores"]
            if errors:
                error_text = "\n".join(f"{item['Stack_ID']}: {item['Error']}" for item in errors[:5])
                messagebox.showwarning(
                    "Conteos actualizados con advertencias",
                    f"Procesados: {result['Procesados']}\nCambios: {result['Cambiados']}\nErrores: {len(errors)}"
                    + (f"\n\nCambios de ejemplo:\n{detail}" if detail else "")
                    + f"\n\nPrimeros errores:\n{error_text}",
                )
            else:
                messagebox.showinfo(
                    "Objetos válidos actualizados",
                    f"Procesados: {result['Procesados']}\nCambios: {result['Cambiados']}"
                    + (f"\n\nCambios de ejemplo:\n{detail}" if detail else "")
                    + f"\n\nSe regeneró {COUNTS_NAME}.",
                )
        except Exception as exc:
            messagebox.showerror("No se pudieron actualizar los conteos", str(exc))

    def _staging_is_frozen(self):
        if self.staging.empty or "Experimento" not in self.staging:
            return False
        experiments = {str(value).strip() for value in self.staging["Experimento"] if str(value).strip()}
        return bool(experiments) and experiments.issubset(FROZEN_EXPERIMENTS)

    def _show_frozen_notice(self):
        return messagebox.showinfo(
            "Experimento maestro bloqueado",
            "HB_02, SCOP_LiCl_02 y SCOP_LiCl_03 están congelados en la V1.2. "
            "No se modificarán sus ROI, metadatos, volúmenes ni orientación.\n\n"
            "Puedes revisar sus mosaicos y ejecutar nuevamente el análisis completo.",
        )

    def _refresh_project(self):
        self.data = load_registry(PROJECT)
        if not hasattr(self, "project_tree"):
            return
        self.project_tree.delete(*self.project_tree.get_children())
        for index, row in self.data.reset_index(drop=True).iterrows():
            values = []
            for column in PROJECT_COLUMNS:
                if column == "Incluir_analisis":
                    values.append("☑" if selected_for_analysis(row.get(column, "")) else "☐")
                else:
                    values.append(row.get(column, ""))
            self.project_tree.insert(
                "", "end", iid=str(index), values=values
            )
        self.experiment_tree.delete(*self.experiment_tree.get_children())
        summary = experiment_summary(self.data)
        for index, row in summary.iterrows():
            self.experiment_tree.insert("", "end", iid=str(index), values=list(row))
        selected = self.data["Incluir_analisis"].map(selected_for_analysis) if not self.data.empty else pd.Series(dtype=bool)
        pending = int((self.data["QC"].eq("Revisar") & selected).sum()) if not self.data.empty else 0
        ok = int((self.data["QC"].eq("OK") & selected).sum()) if not self.data.empty else 0
        selected_animals = (
            self.data.loc[selected, ["Experimento", "Animal"]].drop_duplicates().shape[0]
            if not self.data.empty else 0
        )
        self.project_status.set(
            f"{len(summary)} experimentos · {selected_animals} animales seleccionados · "
            f"{ok} stacks QC OK · {pending} seleccionados por revisar · Datos: {PROJECT}"
        )

    def toggle_analysis_cell(self, event):
        row_id = self.project_tree.identify_row(event.y)
        column_id = self.project_tree.identify_column(event.x)
        if not row_id or not column_id:
            return
        column_index = int(column_id.removeprefix("#")) - 1
        if column_index < 0 or PROJECT_COLUMNS[column_index] != "Incluir_analisis":
            return
        self._toggle_analysis_row(row_id)

    def toggle_selected_animal(self, _event=None):
        selected = self.project_tree.selection()
        if selected:
            self._toggle_analysis_row(selected[0])
        return "break"

    def _toggle_analysis_row(self, row_id):
        try:
            row = self.data.reset_index(drop=True).iloc[int(row_id)]
            experiment = str(row["Experimento"])
            animal = str(row["Animal"])
            include = not selected_for_analysis(row.get("Incluir_analisis", ""))
            result = set_animal_analysis_selection(
                PROJECT, experiment, animal, include
            )
            self.data = result["Data"]
            self._refresh_project()
            matches = self.data.index[
                self.data["Experimento"].eq(experiment)
                & self.data["Animal"].eq(animal)
            ]
            if len(matches):
                self.project_tree.selection_set(str(int(matches[0])))
        except Exception as exc:
            messagebox.showerror("No se pudo actualizar la selección", str(exc))

    def load_saved_experiment(self):
        selected = self.experiment_tree.selection()
        if not selected:
            return messagebox.showinfo("Sin experimento", "Selecciona un experimento guardado.")
        values = self.experiment_tree.item(selected[0], "values")
        experiment = str(values[0])
        self.data = update_saved_bregmas_from_catalog(PROJECT, self.bregmas)["Data"]
        self.staging = self.data[self.data["Experimento"].eq(experiment)].copy().reset_index(drop=True)
        self._reapply_catalogs_to_staging("bregma")
        self.detected_experiment.set(experiment)
        self.scan_status.set(
            f"Experimento guardado {experiment} cargado para edición; "
            "se aplicó el catálogo actual y se conservaron las correcciones manuales."
        )
        self._refresh_staging()
        self.notebook.select(self.load_tab)
        self.open_roi_orientation_png()

    def delete_saved_experiment(self):
        selected = self.experiment_tree.selection()
        if not selected:
            return messagebox.showinfo("Sin experimento", "Selecciona un experimento guardado.")
        values = self.experiment_tree.item(selected[0], "values")
        experiment = str(values[0])
        stacks = values[1]
        animals = values[2]
        answer = messagebox.askyesno(
            "Eliminar experimento del proyecto activo",
            f"Se retirará {experiment} ({stacks} stacks, {animals} animales) del registro activo.\n\n"
            "Sus archivos se moverán a experimentos_eliminados y podrán recuperarse manualmente. "
            "Las corridas históricas no se modificarán.\n\n¿Continuar?",
        )
        if not answer:
            return
        try:
            result = archive_experiment(PROJECT, experiment)
            self.data = result["Data"]
            active_staging = {
                str(value).strip().casefold()
                for value in self.staging.get("Experimento", pd.Series(dtype=str))
                if str(value).strip()
            }
            if experiment.casefold() in active_staging:
                self.clear_staging()
            self._refresh_project()
            messagebox.showinfo(
                "Experimento eliminado",
                f"{result['Experimento']} se retiró del proyecto activo.\n\n"
                f"Respaldo recuperable:\n{result['Archivo']}",
            )
        except Exception as exc:
            messagebox.showerror("No se pudo eliminar el experimento", str(exc))

    def run(self):
        if is_project_inside_app(PROJECT):
            return messagebox.showerror(
                "Proyecto ligado a una app anterior",
                "El análisis no se ejecutará dentro de una carpeta versionada de la aplicación.\n\n"
                "Usa «Recuperar / cambiar proyecto» y migra el proyecto a la carpeta "
                "proyecto_cfos estable.",
            )
        self.data = load_registry(PROJECT)
        if self.data.empty:
            return messagebox.showerror("Sin datos", "Primero guarda al menos un experimento.")
        selected = self.data["Incluir_analisis"].map(selected_for_analysis)
        pending = self.data[self.data["QC"].eq("Revisar") & selected]
        if not pending.empty:
            identifiers = ", ".join(pending["Stack_ID"].head(8))
            suffix = "…" if len(pending) > 8 else ""
            return messagebox.showerror(
                "Proyecto pendiente de revisión",
                f"Hay {len(pending)} stack(s) con QC=Revisar: {identifiers}{suffix}\n\n"
                "Corrígelos o márcalos como Excluir antes del análisis final.",
            )
        ok = self.data[self.data["QC"].eq("OK") & selected]
        if ok.empty:
            return messagebox.showerror(
                "Sin stacks válidos",
                "No hay animales seleccionados con stacks QC=OK.",
            )
        experiments = ok["Experimento"].nunique()
        animals = ok[["Experimento", "Animal"]].drop_duplicates().shape[0]
        answer = messagebox.askyesno(
            "Ejecutar análisis completo",
            f"Se analizarán {len(ok)} stacks de {animals} animales seleccionados "
            f"en {experiments} experimento(s).\n\n"
            "¿Continuar?",
        )
        if not answer:
            return
        try:
            count_file = write_counts_workbook(PROJECT, self.data)
            output = run_final(PROJECT)
            shutil.copy2(count_file, output / COUNTS_NAME)
            warning_file = output / "ADVERTENCIA_stacks_con_error.txt"
            if warning_file.exists():
                messagebox.showwarning(
                    "Análisis terminado con advertencias",
                    f"La corrida se generó, pero uno o más stacks fallaron. "
                    f"Revisa la hoja Errores y {warning_file.name}.\n\n{output}",
                )
            else:
                messagebox.showinfo(
                    "Análisis terminado",
                    f"La corrida final y una copia de {COUNTS_NAME} están en:\n{output}",
                )
        except Exception as exc:
            messagebox.showerror("Error durante el análisis", str(exc))


if __name__ == "__main__":
    App().mainloop()
