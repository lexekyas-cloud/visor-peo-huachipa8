import os, sys
from pathlib import Path
from dotenv import load_dotenv

def fail(msg):
    print(f"[X] {msg}")
    sys.exit(1)

load_dotenv()

excel_path = os.getenv("EXCEL_PATH")
sheet = os.getenv("EXCEL_SHEET")
upload_dir = os.getenv("UPLOAD_DIR")

if not excel_path:
    fail("Falta EXCEL_PATH en .env")
p = Path(excel_path)
if not p.exists():
    fail(f"No existe el archivo: {p}")

if sheet is None or sheet == "":
    fail("Falta EXCEL_SHEET en .env (usa 0 o nombre de hoja)")

if upload_dir:
    u = Path(upload_dir)
    if not u.exists():
        print(f"[!] La carpeta UPLOAD_DIR no existe: {u}")
        print("    Créala para evitar errores al subir evidencias.")
else:
    print("[!] No se definió UPLOAD_DIR (opcional).")

print("[✓] .env válido. Ruta Excel OK.")