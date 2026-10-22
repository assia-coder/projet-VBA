"""
excel_bridge.py - Point d'entrée appelé par la macro VBA Excel.

Usage :
    python excel_bridge.py passive   -> train_passive_ml.py
    python excel_bridge.py active    -> train_active_ml.py
    python excel_bridge.py prophet   -> train_prophet.py
    python excel_bridge.py predict   -> predict_next_30_days.py (modèle gagnant, J+1..J+30)
    python excel_bridge.py workbook  -> generate_workbook.py
    python excel_bridge.py monthly   -> monthly_forecast.py + generate_workbook.py (hors flux principal)
    python excel_bridge.py all       -> preprocess, 3 entraînements, compare_all (winner.json),
                                        predict_next_30_days, generate_workbook

Écrit output/status.txt, lu par la macro VBA :
    ligne 1 : RUNNING | SUCCESS | ERROR
    ligne 2 : message
    ligne 3 : horodatage
    ligne 4 : durée en secondes
Le journal complet de l'exécution est dans output/bridge_log.txt.
"""
import os
import subprocess
import sys
import time
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
STATUS_FILE = os.path.join(OUTPUT_DIR, "status.txt")
LOG_FILE = os.path.join(OUTPUT_DIR, "bridge_log.txt")
LONG_FILE = os.path.join(OUTPUT_DIR, "series_long.csv")

# Scripts lancés pour chaque commande
COMMANDS = {
    "passive": ["train_passive_ml.py"],
    "active": ["train_active_ml.py"],
    "prophet": ["train_prophet.py"],
    # Prévision J+1..J+30 avec le modèle gagnant (winner.json)
    "predict": ["predict_next_30_days.py"],
    # Régénère Application_Prevision.xlsx et les CSV lus par les UserForms
    "workbook": ["generate_workbook.py"],
    # Prévision mensuelle M+1..M+3 : hors flux principal, lancée à la demande
    "monthly": ["monthly_forecast.py", "generate_workbook.py"],
    "all": [
        "preprocess.py",
        "train_passive_ml.py",
        "train_active_ml.py",
        "train_prophet.py",
        "compare_all.py",            # identifie le gagnant + écrit winner.json
        "predict_next_30_days.py",   # prédit avec le gagnant
        "generate_workbook.py",      # génère Application_Prevision.xlsx
    ],
}

# Encodage console : évite les erreurs d'accents sous Windows
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


def write_status(status, message, start=None):
    """Écrit status.txt en ANSI (cp1252) pour une lecture directe par VBA."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    duration = f"{time.time() - start:.0f}" if start else "0"
    content = "\n".join([status, message, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), duration])
    with open(STATUS_FILE, "w", encoding="cp1252", errors="replace") as f:
        f.write(content + "\n")


def run_script(script, log):
    """Lance un script Python, affiche et journalise sa sortie. Retourne le code de sortie."""
    path = os.path.join(BASE_DIR, script)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Script introuvable : {script}")

    banner = f"\n##### {datetime.now():%H:%M:%S} - lancement de {script} #####\n"
    print(banner, flush=True)
    log.write(banner)

    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
    proc = subprocess.Popen([sys.executable, path], cwd=BASE_DIR, env=env,
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            text=True, encoding="utf-8", errors="replace")
    last_lines = []
    for line in proc.stdout:
        print(line, end="", flush=True)
        log.write(line)
        last_lines = (last_lines + [line.strip()])[-80:]
    proc.wait()
    return proc.returncode, last_lines


def error_message(script, last_lines):
    """Extrait le message d'erreur le plus parlant de la sortie du script."""
    for line in reversed(last_lines):
        if line.startswith("[ERREUR]"):
            return f"{script} : {line}"
    tail = next((l for l in reversed(last_lines) if l), "erreur inconnue")
    return f"{script} : {tail}"


def main():
    # Vérification de l'argument
    if len(sys.argv) != 2 or sys.argv[1].lower() not in COMMANDS:
        msg = f"Argument invalide. Usage : python excel_bridge.py [{' | '.join(COMMANDS)}]"
        print(msg)
        write_status("ERROR", msg)
        return 1
    cmd = sys.argv[1].lower()
    start = time.time()

    # Le preprocessing doit avoir été fait (sauf si la commande le lance elle-même)
    if "preprocess.py" not in COMMANDS[cmd] and not os.path.exists(LONG_FILE):
        msg = "series_long.csv introuvable : lancez d'abord python preprocess.py"
        print(f"[ERREUR] {msg}")
        write_status("ERROR", msg, start)
        return 1

    write_status("RUNNING", f"Commande '{cmd}' en cours...", start)
    scripts = COMMANDS[cmd]

    with open(LOG_FILE, "w", encoding="utf-8") as log:
        log.write(f"excel_bridge.py {cmd} - {datetime.now():%Y-%m-%d %H:%M:%S}\n")
        for i, script in enumerate(scripts, start=1):
            write_status("RUNNING", f"Etape {i}/{len(scripts)} : {script}", start)
            code, last_lines = run_script(script, log)
            if code != 0:
                msg = error_message(script, last_lines)
                print(f"\n[ERREUR] Échec de {script} (code {code}). Pipeline arrêté.")
                write_status("ERROR", msg, start)
                return 1

    msg = f"Commande '{cmd}' terminee : {', '.join(scripts)}"
    write_status("SUCCESS", msg, start)
    print(f"\n[OK] {msg} en {time.time() - start:.0f}s")
    print(f"     Statut écrit dans {STATUS_FILE}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print(f"[ERREUR] excel_bridge : {e}")
        write_status("ERROR", f"excel_bridge : {e}")
        sys.exit(1)
