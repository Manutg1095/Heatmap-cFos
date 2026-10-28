#!/bin/zsh
set -u

script_dir="${0:A:h}"
cd "$script_dir" || exit 1

if command -v conda >/dev/null 2>&1; then
  eval "$(conda shell.zsh hook)"
  conda activate cfos 2>/dev/null || true
elif [[ -x "$HOME/miniconda3/bin/conda" ]]; then
  eval "$("$HOME/miniconda3/bin/conda" shell.zsh hook)"
  conda activate cfos 2>/dev/null || true
elif [[ -x "$HOME/anaconda3/bin/conda" ]]; then
  eval "$("$HOME/anaconda3/bin/conda" shell.zsh hook)"
  conda activate cfos 2>/dev/null || true
fi

if command -v python >/dev/null 2>&1; then
  python_bin="python"
elif command -v python3 >/dev/null 2>&1; then
  python_bin="python3"
else
  print "No se encontró Python. Activa el entorno cfos y vuelve a abrir este archivo."
  read -k 1 "?Pulsa una tecla para cerrar."
  exit 1
fi

"$python_bin" iniciar.py
heatmap_exit_code=$?
if [[ $heatmap_exit_code -ne 0 ]]; then
  print "Heatmap V1.2 terminó con código $heatmap_exit_code."
  read -k 1 "?Pulsa una tecla para cerrar."
fi
exit $heatmap_exit_code
