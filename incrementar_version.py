import re
import sys

file_path = "auditor_facturas.py"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Buscar la versión actual
match = re.search(r'"version":\s*"2\.0(?:\.(\d+))?"', content)
if not match:
    print("No se encontró la versión base en el archivo.")
    sys.exit(1)

current_patch = match.group(1)
new_patch = 1 if current_patch is None else int(current_patch) + 1

old_version = f"2.0.{current_patch}" if current_patch else "2.0"
new_version = f"2.0.{new_patch}"

# Realizar los reemplazos
content = content.replace(f'"version": "{old_version}"', f'"version": "{new_version}"')
content = content.replace(f'(v{old_version})', f'(v{new_version})')
content = content.replace(f'Versión: {old_version}\\n', f'Versión: {new_version}\\n')

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print(f"Versión de Auditor de Facturas actualizada a: {new_version}")
