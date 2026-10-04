import os
import shutil
from pathlib import Path

project_root = Path(__file__).resolve().parent.parent
dest_dir = project_root / "Tesseract-OCR"
dest_dir.mkdir(exist_ok=True)

# Archivos y carpetas de Tesseract a mover
items_to_move = [
    "tessdata", "doc", "tesseract.exe", "tesseract-uninstall.exe", 
    "ambiguous_words.exe", "classifier_tester.exe", "cntraining.exe", 
    "combine_lang_model.exe", "combine_tessdata.exe", "dawg2wordlist.exe", 
    "lstmeval.exe", "lstmtraining.exe", "merge_unicharsets.exe", 
    "mftraining.exe", "set_unicharset_properties.exe", "shapeclustering.exe", 
    "text2image.exe", "unicharset_extractor.exe", "winpath.exe", "wordlist2dawg.exe"
]

# Mover también todos los archivos .dll y .html
for file in os.listdir(project_root):
    if file.endswith('.dll') or file.endswith('.html'):
        if file not in items_to_move:
            items_to_move.append(file)

for item in items_to_move:
    source = project_root / item
    if source.exists():
        shutil.move(str(source), str(dest_dir / item))
        print(f"Movido: {item}")

print("\n¡Listo! Todos los archivos de Tesseract han sido organizados.")
