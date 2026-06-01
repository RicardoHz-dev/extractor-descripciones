from pathlib import Path
from datetime import datetime

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

LOG_GENERAL = LOG_DIR / f"log_{timestamp}.txt"
LOG_DETALLE = LOG_DIR / f"detalle_{timestamp}.txt"

def escribir_log(mensaje):
    with open(LOG_GENERAL, "a", encoding="utf-8") as f:
        f.write(str(mensaje) + "\n")

def escribir_detalle(mensaje):
    with open(LOG_DETALLE, "a", encoding="utf-8") as f:
        f.write(str(mensaje) + "\n")