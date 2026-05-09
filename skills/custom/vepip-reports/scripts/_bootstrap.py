"""Ensure required Python packages are importable. Pip-installs if missing.

When the deer-flow harness is properly synced (see backend/packages/harness/
pyproject.toml), python-pptx / python-docx / reportlab / Pillow are already
present and this is a no-op. The fallback install is defensive insurance for
environments where uv sync hasn't been run yet.
"""
import importlib
import importlib.util
import shutil
import subprocess
import sys


def ensure(packages: dict[str, str]) -> None:
    """packages: {import_name: pip_name}. Installs missing pip_names quietly."""
    missing = [pip_name for import_name, pip_name in packages.items()
               if importlib.util.find_spec(import_name) is None]
    if not missing:
        return

    print(f"[bootstrap] installing: {' '.join(missing)}", flush=True)

    # Prefer `uv pip install` since the deer-flow venv is uv-managed and
    # `pip` itself may not be present after `uv sync`.
    uv = shutil.which("uv")
    if uv:
        try:
            subprocess.check_call(
                [uv, "pip", "install", "--quiet", "--python", sys.executable, *missing]
            )
            return
        except subprocess.CalledProcessError as e:
            print(f"[bootstrap] uv pip install failed ({e}); falling back to pip", flush=True)

    # Fall back to python -m pip if uv path didn't work.
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check", *missing]
        )
    except subprocess.CalledProcessError as e:
        raise SystemExit(
            f"[bootstrap] failed to install {missing}. "
            f"Run `uv sync` in deer-flow/backend, then retry. ({e})"
        ) from e
