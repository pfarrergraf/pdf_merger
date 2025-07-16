# email_utils.py

import os
import pythoncom
import win32com.client

def send_email_via_outlook(
    subject: str,
    body: str,
    to: str,
    attachments: list[str] = None,
    display_before_send: bool = False
) -> None:
    """
    Uses Outlook COM to create/send an email. Initializes COM for this thread.
    
    :param subject:           Subject line
    :param body:              Plain‐text body
    :param to:                Semicolon‐separated recipient(s), e.g. "me@domain.com"
    :param attachments:       Optional list of file paths to attach
    :param display_before_send: if True, opens the draft so you can review & click Send
    """
    # Initialize COM in this thread
    pythoncom.CoInitialize()
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail    = outlook.CreateItem(0)  # olMailItem == 0

        # Fill in fields
        mail.Subject = subject
        mail.Body    = body
        mail.To      = to

        # Add attachments if any
        if attachments:
            for path in attachments:
                full_path = os.path.abspath(path)
                mail.Attachments.Add(full_path)

        # Send or display
        if display_before_send:
            mail.Display()   # show draft for manual Send
        else:
            mail.Send()      # send immediately

        print("[✉️] Email sent via Outlook!")
    finally:
        # Always uninitialize COM on this thread when done
        pythoncom.CoUninitialize()
