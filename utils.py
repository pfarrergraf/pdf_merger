# utils.py

import os
import pythoncom
from docx2pdf import convert as docx_to_pdf
import win32com.client
from PyPDF2 import PdfMerger

def convert_docx(input_path: str, output_path: str) -> bool:
    """
    Convert a .docx to .pdf via Word COM. Returns True on success.
    """
    pythoncom.CoInitialize()      # COM initialisieren
    try:
        # docx2pdf nutzt Word COM unter der Haube
        docx_to_pdf(input_path, output_path)
        return True
    except Exception as e:
        print(f"[!] DOCX→PDF failed: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()  # COM wieder freigeben

def convert_xlsx(input_path: str, output_path: str) -> bool:
    """
    Convert .xlsx to .pdf via Excel COM. Returns True on success.
    """
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(os.path.abspath(input_path))
        wb.ExportAsFixedFormat(0, os.path.abspath(output_path))
        wb.Close(False)
        excel.Quit()
        return True
    except Exception as e:
        print(f"[!] XLSX→PDF failed: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

def merge_pdfs(pdf_paths: list[str], output_path: str) -> bool:
    """
    Merge a list of PDFs into a single PDF at output_path.
    """
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

def dummy_convert(input_path: str, output_path: str) -> bool:
    """
    Dummy-Converter: Erstellt eine minimale, leere PDF-Datei.
    Damit kannst du Merge- und Email-Logik testen, ohne Word/Excel.
    """
    try:
        # Eine sehr einfache PDF-Grundstruktur
        pdf_content = (
            b"%PDF-1.4\n"
            b"%\xe2\xe3\xcf\xd3\n"
            b"1 0 obj<<>>endobj\n"
            b"trailer<<>>\n"
            b"%%EOF\n"
        )
        with open(output_path, "wb") as f:
            f.write(pdf_content)
        return True
    except Exception as e:
        print(f"[!] dummy_convert failed: {e}")
        return False
