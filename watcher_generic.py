# watcher_generic.py

import os
import sys
import json
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Echte Converter (Word/Excel)
from utils import convert_docx, convert_xlsx, merge_pdfs

# E-Mail via Outlook
from email_utils import send_email_via_outlook

# Konfiguration
from config import WATCH_FOLDER, EMAIL_TO


def load_client_config(path="config_clients.json") -> dict:
    """
    Lädt die Mandanten-Konfiguration aus einer JSON-Datei.
    """
    if not os.path.isfile(path):
        print(f"[!] Konfigurationsdatei fehlt: {path}")
        sys.exit(1)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


class GenericHandler(FileSystemEventHandler):
    """
    Beobachtet WATCH_FOLDER und verarbeitet Dateien nach Client-Config:
    - Konvertiert docx/xlsx → pdf
    - Mergt, wenn alle für einen Mandanten vorhanden sind
    - Sendet Outlook-Mail mit dem gemergten PDF
    """
    def __init__(self, clients_config: dict):
        super().__init__()
        self.clients = clients_config
        # Status: client_id → Set der erzeugten PDF-Namen
        self.status = {cid: set() for cid in self.clients}

    def on_created(self, event):
        # 0) Ignoriere Ordner-Events
        if event.is_directory:
            return

        fname = os.path.basename(event.src_path)
        root, ext = os.path.splitext(fname)

        # 1) Ignoriere temporäre Office-Dateien und .tmp
        if fname.startswith("~$") or ext.lower() == ".tmp":
            return

        # 2) Ignoriere Merge-PDF selbst (endet auf _Haushalt_<Jahr>.pdf)
        for cid, info in self.clients.items():
            merged_name = f"{cid}_Haushalt_{info['year']}.pdf"
            if fname == merged_name:
                return

        print(f"[+] Neue Datei entdeckt: {fname}")

        # 3) Prüfe alle Mandanten auf Datei-Patterns
        for cid, info in self.clients.items():
            for file_def in info["files"]:
                pattern = file_def["pattern"]
                action  = file_def["convert"]

                # DEBUG-Print (optional)
                print(f"    [DEBUG] Prüfe Mandant {cid}: pattern={pattern}, action={action}")

                # 4) Wenn der Dateiname exakt passt, konvertiere
                if fname == pattern:
                    pdf_name = os.path.splitext(pattern)[0] + ".pdf"
                    pdf_path = os.path.join(WATCH_FOLDER, pdf_name)

                    success = False
                    if action == "docx":
                        success = convert_docx(event.src_path, pdf_path)
                    elif action == "xlsx":
                        success = convert_xlsx(event.src_path, pdf_path)

                    if success:
                        print(f"    ✓ {cid}: {pattern} → {pdf_name}")
                        self.status[cid].add(pdf_name)
                    # Nach Bearbeitung eines Patterns abbrechen
                    break

        # 5) Prüfe, ob ein Mandant komplett ist
        for cid, done_pdfs in list(self.status.items()):
            expected = {
                os.path.splitext(fd["pattern"])[0] + ".pdf"
                for fd in self.clients[cid]["files"]
            }
            if done_pdfs == expected:
                print(f"[✔] Mandant {cid} vollständig – merge & email")

                # 6) Merge in konfigurierter Reihenfolge
                order = self.clients[cid]["merge_order"]
                pdf_paths = [os.path.join(WATCH_FOLDER, name) for name in order]
                merged_name = f"{cid}_Haushalt_{self.clients[cid]['year']}.pdf"
                merged_path = os.path.join(WATCH_FOLDER, merged_name)

                if merge_pdfs(pdf_paths, merged_path):
                    print(f"    ✔ Gemergt: {merged_path}")
                    # 7) E-Mail via Outlook senden
                    send_email_via_outlook(
                        subject=f"Mandant {cid}: Haushaltsentwurf fertig",
                        body=f"Der Haushaltsentwurf wurde erstellt:\n{merged_name}",
                        to=EMAIL_TO,
                        attachments=[merged_path],
                        display_before_send=True
                    )

                # 8) Mandanten-Status löschen, damit nicht erneut getriggert wird
                del self.status[cid]
                break


if __name__ == "__main__":
    # A) Prüfungen vor dem Start
    if not os.path.isdir(WATCH_FOLDER):
        print(f"[!] WATCH_FOLDER existiert nicht: {WATCH_FOLDER}")
        sys.exit(1)
    if not os.path.isfile("config_clients.json"):
        print("[!] config_clients.json fehlt im Arbeitsverzeichnis.")
        sys.exit(1)

    # B) Konfiguration laden & Anzeigen
    clients = load_client_config("config_clients.json")
    print(f"[>] Geladene Mandanten: {list(clients.keys())}")
    for cid, info in clients.items():
        print(f"    Mandant {cid}: {info['name']}")

    # C) Watchdog-Observer initialisieren
    handler  = GenericHandler(clients)
    observer = Observer()
    observer.schedule(handler, WATCH_FOLDER, recursive=False)
    observer.start()
    print(f"[>] Überwache Ordner: {WATCH_FOLDER}")

    # D) Loop bis Strg+C
    try:
        while observer.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
