"""Ensure required Python packages are importable. Pip-installs if missing.

Used by build_pptx / build_docx / build_pdf so the skill works in any
sandbox without requiring the base image to ship with these packages.
"""
import importlib
import subprocess
import sys


def ensure(packages: dict[str, str]) -> None:
    """packages: {import_name: pip_name}. Installs missing pip_names quietly."""
    missing = []
    for import_name, pip_name in packages.items():
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing.append(pip_name)
    if not missing:
        return
    print(f"[bootstrap] installing: {' '.join(missing)}", flush=True)
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check", *missing]
    )
