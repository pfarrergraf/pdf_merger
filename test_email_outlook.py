# test_email_outlook.py
# test_email_outlook.py
from email_utils import send_email_via_outlook
from config import EMAIL_TO

if __name__ == "__main__":
    send_email_via_outlook(
        subject="[Test] Outlook COM Email",
        body="✅ If you see this in Outlook Draft (and/or Inbox), it works!",
        to=EMAIL_TO,
        display_before_send=True
    )