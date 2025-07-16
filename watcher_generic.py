# watcher_generic.py

import os
import sys
import json
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Wir nutzen unsere bestehenden Helfer:
from utils import convert_docx, convert_xlsx, merge_pdfs
from email_utils import send_email_via_outlook
from config import WATCH_FOLDER, EMAIL_TO

# ──────────────────────────────────────────────────────────────────────────────
# 1) Konfigurations-Loader
#    Wir lesen eine JSON-Datei mit Mandanten-Infos ein, z.B.:
#
#    {
#      "6840": {
#        "name": "Kirchengemeinde Oberlahnstein",
#        "year": 2025,
#        "files": [
#          { "pattern": "6840_Budget.xlsx", "convert": "xlsx" },
#          { "pattern": "6840_Entwurf.docx", "convert": "docx" }
#        ],
#        "merge_order": ["6840_Entwurf.pdf", "6840_Budget.pdf"]
#      },
#      ...
#    }
# ──────────────────────────────────────────────────────────────────────────────

def load_client_config(path="config_clients.json") -> dict:
    """Lädt die Mandanten-Konfiguration aus einer JSON-Datei."""
    if not os.path.isfile(path):
        print(f"[!] Konfigurationsdatei fehlt: {path}")
        sys.exit(1)
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data


# ──────────────────────────────────────────────────────────────────────────────
# 2) Handler-Klasse: Reagiert auf neue Dateien im Watch-Ordner
#    Wir behalten für jeden Mandanten den Status, welche PDFs schon fertig sind.
# ──────────────────────────────────────────────────────────────────────────────

class GenericHandler(FileSystemEventHandler):
    def __init__(self, clients_config):
        super().__init__()
        self.clients = clients_config
        # Status-Dict: client_id → set of fertiger PDF-Dateinamen
        self.status = {cid: set() for cid in self.clients}
    
    def on_created(self, event):
        """Wird aufgerufen, wenn eine Datei ankommt."""
        fname = os.path.basename(event.src_path)
        # 1) Ignoriere temporäre Office-Dateien:
        if fname.startswith("~$"):
            return
        
        print(f"[+] Neue Datei: {fname}")
        
        # 2) Für jeden Mandanten prüfen, ob der Dateiname passt
        for cid, info in self.clients.items():
            for file_def in info["files"]:
                pattern = file_def["pattern"]
                action  = file_def["convert"]  # "docx" oder "xlsx"
                
                # Passt der Name genau?
                if fname == pattern:
                    # Berechne den Ziel-PDF-Pfad im Watch-Ordner
                    pdf_name = os.path.splitext(pattern)[0] + ".pdf"
                    pdf_path = os.path.join(WATCH_FOLDER, pdf_name)
                    
                    # 3) Konvertierung aufrufen
                    if action == "docx":
                        success = convert_docx(event.src_path, pdf_path)
                    elif action == "xlsx":
                        success = convert_xlsx(event.src_path, pdf_path)
                    else:
                        success = False
                    
                    if success:
                        print(f"[✓] {cid}: {pattern} → {pdf_name}")
                        # 4) Status merken
                        self.status[cid].add(pdf_name)
                    
                    # Nach der Konvertierung brauchen wir nicht weiter prüfen
                    break
        
        # 5) Nachsehen, ob bei einem Mandanten alle PDFs fertig sind
        for cid, done_pdfs in self.status.items():
            expected = { os.path.splitext(fd["pattern"])[0] + ".pdf"
                         for fd in self.clients[cid]["files"] }
            if done_pdfs == expected:
                print(f"[✔] Mandant {cid} komplett: Merge & Mail")
                
                # 6) Merge in der in der Config angegebenen Reihenfolge
                order = self.clients[cid]["merge_order"]
                pdf_paths = [os.path.join(WATCH_FOLDER, name) for name in order]
                merged_name = f"{cid}_Haushalt_{self.clients[cid]['year']}.pdf"
                merged_path = os.path.join(WATCH_FOLDER, merged_name)
                
                if merge_pdfs(pdf_paths, merged_path):
                    # 7) Benachrichtigung
                    send_email_via_outlook(
                        subject=f"Mandant {cid}: Haushaltsentwurf fertig",
                        body=f"Der Haushaltsentwurf {merged_name} wurde erstellt.",
                        to=EMAIL_TO,
                        attachments=[merged_path],
                        display_before_send=True
                    )
                # 8) Einmal pro Mandant stoppen wir die Überwachung oder resetten den Status
                #    (hier Beispiel: nur diesen Mandanten nicht mehr beachten)
                del self.status[cid]
                break  # nur einen Mandanten pro Event bearbeiten


# ──────────────────────────────────────────────────────────────────────────────
# 3) Main: Watcher starten
# ──────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # 0) Prüfungen
    if not os.path.isdir(WATCH_FOLDER):
        print(f"[!] WATCH_FOLDER existiert nicht: {WATCH_FOLDER}")
        sys.exit(1)
    if not os.path.exists("config_clients.json"):
        print("[!] Bitte erstelle config_clients.json im aktuellen Ordner.")
        sys.exit(1)
    
    # 1) Config laden
    clients = load_client_config("config_clients.json")
    print(f"[>] Geladene Mandanten: {list(clients.keys())}")
    
    for cid, info in clients.items():
        print(f"  Mandant {cid}: {info['name']}")

    # 2) Watcher initialisieren
    handler  = GenericHandler(clients)
    observer = Observer()
    observer.schedule(handler, WATCH_FOLDER, recursive=False)
    observer.start()
    print(f"[>] Watching folder: {WATCH_FOLDER}")
    
    # 3) Loop
    try:
        while observer.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
