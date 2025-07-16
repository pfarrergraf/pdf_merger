# utils.py
import pythoncom
import os
from docx2pdf import convert as docx_to_pdf
import win32com.client
from PyPDF2 import PdfMerger

# email
import smtplib
from email.message import EmailMessage

def convert_docx(input_path: str, output_path: str) -> bool:
    """Convert a .docx to .pdf. Returns True on success."""
# Initialize COM in this thread
    pythoncom.CoInitialize()
    try:
        docx_to_pdf(input_path, output_path)
        return True
    except Exception as e:
        print(f"[!] DOCX→PDF failed: {e}")
        return False
    finally:
        pythoncom.CoInitialize()


def convert_xlsx(input_path: str, output_path: str) -> bool:
    """Convert .xlsx to .pdf via Excel COM."""
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(os.path.abspath(input_path))
        wb.ExportAsFixedFormat(0, os.path.abspath(output_path))  # 0=PDF
        wb.Close(False)
        excel.Quit()
        return True
    except Exception as e:
        print(f"[!] XLSX→PDF failed: {e}")
        return False
    finally:
        pythoncom.CoInitialize()


def merge_pdfs(pdf_paths: list[str], output_path: str) -> bool:
    """Merge list of PDFs in order, write to output_path."""
    try:
        merger = PdfMerger()
        for p in pdf_paths:
            merger.append(p)
        merger.write(output_path)
        merger.close()
        return True
    except Exception as e:
        print(f"[!] PDF merge failed: {e}")
        return False



def send_email(subject: str, body: str,
               from_addr: str, to_addr: str,
               smtp_server: str, smtp_port: int,
               username: str, password: str):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg.set_content(body)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(username, password)
        server.send_message(msg)
        print("[✉️] Email sent!")
