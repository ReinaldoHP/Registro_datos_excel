"""Incrementa el número de parche (MAJOR.MINOR.PATCH) del Auditor de Facturas.

Actualiza de forma coherente:
  - auditor_facturas.py   ("version": "X.Y.Z", (vX.Y.Z), Versión: X.Y.Z\\n)
  - auditor_facturas.spec (name='auditor_facturas vX.Y.Z')
  - abrir_auditor.bat     (dist\\auditor_facturas vX.Y.Z.exe)

Uso:
  python scripts/incrementar_version.py [--root RUTA]

La raíz del proyecto puede indicarse con --root o con la variable de entorno
AUDITOR_ROOT (útil para probar sobre copias). Si alguno de los textos esperados
no se encuentra, el script termina con error sin modificar ningún archivo.
"""
import argparse
import os
import re
import sys
from pathlib import Path

VERSION_RE = re.compile(r'"version":\s*"(\d+)\.(\d+)(?:\.(\d+))?"')


def leer(path: Path) -> str:
    # newline="" conserva los finales de línea originales (CRLF en .bat)
    with open(path, "r", encoding="utf-8", newline="") as f:
        return f.read()


def escribir(path: Path, content: str) -> None:
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(content)


def reemplazar(content: str, old: str, new: str, path: Path, errores: list) -> str:
    if old not in content:
        errores.append(f"{path.name}: no se encontró el texto {old!r}")
        return content
    return content.replace(old, new)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--root",
        default=os.environ.get("AUDITOR_ROOT") or Path(__file__).resolve().parent.parent,
        help="Raíz del proyecto (por defecto, la carpeta padre de scripts/).",
    )
    args = parser.parse_args(argv)
    root = Path(args.root).resolve()

    py_path = root / "auditor_facturas.py"
    spec_path = root / "auditor_facturas.spec"
    bat_path = root / "abrir_auditor.bat"

    for path in (py_path, spec_path, bat_path):
        if not path.is_file():
            print(f"ERROR: no existe el archivo {path}", file=sys.stderr)
            return 1

    py = leer(py_path)
    match = VERSION_RE.search(py)
    if not match:
        print(f'ERROR: no se encontró "version": "X.Y.Z" en {py_path}', file=sys.stderr)
        return 1

    major, minor, patch = match.group(1), match.group(2), match.group(3)
    old_version = f"{major}.{minor}.{patch}" if patch is not None else f"{major}.{minor}"
    new_version = f"{major}.{minor}.{1 if patch is None else int(patch) + 1}"

    errores = []
    py = reemplazar(py, match.group(0), f'"version": "{new_version}"', py_path, errores)
    py = reemplazar(py, f"(v{old_version})", f"(v{new_version})", py_path, errores)
    py = reemplazar(py, f"Versión: {old_version}\\n", f"Versión: {new_version}\\n", py_path, errores)

    spec = reemplazar(
        leer(spec_path),
        f"name='auditor_facturas v{old_version}'",
        f"name='auditor_facturas v{new_version}'",
        spec_path,
        errores,
    )
    bat = reemplazar(
        leer(bat_path),
        f"dist\\auditor_facturas v{old_version}.exe",
        f"dist\\auditor_facturas v{new_version}.exe",
        bat_path,
        errores,
    )

    if errores:
        print("ERROR: no se actualizó ningún archivo:", file=sys.stderr)
        for e in errores:
            print(f"  - {e}", file=sys.stderr)
        return 1

    escribir(py_path, py)
    escribir(spec_path, spec)
    escribir(bat_path, bat)

    print(f"Versión de Auditor de Facturas actualizada: {old_version} -> {new_version}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
