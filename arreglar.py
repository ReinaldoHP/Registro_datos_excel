import os
import shutil

# Crear la carpeta si no existe
dest_dir = "Tesseract-OCR"
if not os.path.exists(dest_dir):
    os.makedirs(dest_dir)

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
for file in os.listdir('.'):
    if file.endswith('.dll') or file.endswith('.html'):
        if file not in items_to_move:
            items_to_move.append(file)

for item in items_to_move:
    if os.path.exists(item):
        shutil.move(item, os.path.join(dest_dir, item))
        print(f"Movido: {item}")

print("\n¡Listo! Todos los archivos de Tesseract han sido organizados.")
