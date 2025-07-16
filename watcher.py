# watcher.py

import os
import sys
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from config import WATCH_FOLDER, EMAIL_TO
from utils import convert_docx, convert_xlsx, merge_pdfs
from email_utils import send_email_via_outlook

# Die beiden erwarteten PDF-Dateinamen
EXPECTED    = ["a_final.pdf", "b_final.pdf"]
# Name der zusammengeführten Ausgabedatei
MERGED_NAME = "final_package.pdf"

class Handler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self.found = []

    def on_created(self, event):
        """
        Wird aufgerufen, wenn eine Datei erstellt wird.
        Konvertiert a_final.docx → a_final.pdf,
        b_final.xlsx → b_final.pdf,
        und wenn beide PDFs da sind, merged & sendet Mail.
        """
        fname = os.path.basename(event.src_path)

        # 1) Ignoriere temporäre Office-Dateien („~$…“)
        if fname.startswith("~$"):
            return

        print(f"[+] Detected new file: {fname}")

        # 2) Konvertierung
        a_pdf = os.path.join(WATCH_FOLDER, "a_final.pdf")
        b_pdf = os.path.join(WATCH_FOLDER, "b_final.pdf")

        if fname.lower() == "a_final.docx":
            if convert_docx(event.src_path, a_pdf):
                print(f"[+] Converted to {a_pdf}")
                self.found.append("a_final.pdf")

        elif fname.lower() == "b_final.xlsx":
            if convert_xlsx(event.src_path, b_pdf):
                print(f"[+] Converted to {b_pdf}")
                self.found.append("b_final.pdf")

        # 3) Sobald beide PDFs da sind → merge & notify
        if all(name in self.found for name in EXPECTED):
            print("[✔] Both PDFs ready—merging now…")
            ordered = [
                os.path.join(WATCH_FOLDER, "b_final.pdf"),
                os.path.join(WATCH_FOLDER, "a_final.pdf"),
            ]
            merged_path = os.path.join(WATCH_FOLDER, MERGED_NAME)

            if merge_pdfs(ordered, merged_path):
                print(f"[✔] Merged into {merged_path}!")
                notify_user(merged_path)
            observer.stop()  # nur einmal durchlaufen

def notify_user(merged_path: str):
    """
    Sendet per Outlook COM eine E-Mail mit dem gemergten PDF als Anhang.
    """
    send_email_via_outlook(
        subject="✅ Your PDF package is ready",
        body=f"Your merged PDF is ready at:\n{merged_path}",
        to=EMAIL_TO,
        attachments=[merged_path],
        display_before_send=True
    )

if __name__ == "__main__":
    # 0) Existenz-Check des Watch-Ordners
    if not os.path.isdir(WATCH_FOLDER):
        print(f"[!] WATCH_FOLDER existiert nicht: {WATCH_FOLDER}")
        sys.exit(1)

    print(f"[>] Watching folder (absolute): {os.path.abspath(WATCH_FOLDER)}")

    # 4) Observer starten
    event_handler = Handler()
    observer = Observer()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=False)
    observer.start()

    try:
        while observer.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
