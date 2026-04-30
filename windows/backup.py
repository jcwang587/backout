import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils import BASE_DIR, init_db, save_email, get_stats


def backup_inbox():
    try:
        import win32com.client
    except ImportError:
        print("Missing dependency. Run: pip install pywin32")
        return

    print("=== Outlook Email Backup (Windows) ===")
    print("Connecting to Outlook...")

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    total = len(messages)
    print(f"Found {total} emails, saving...")

    con = init_db()
    saved = skipped = already_exists = 0
    for msg in messages:
        try:
            dt = msg.ReceivedTime.replace(tzinfo=None)
            if save_email(
                date_str=dt,
                sender=f"{msg.SenderName} <{msg.SenderEmailAddress}>",
                subject=msg.Subject or "(no subject)",
                body=msg.Body or "",
                con=con,
            ):
                saved += 1
            else:
                already_exists += 1
        except Exception:
            skipped += 1
    con.close()

    print(f"New: {saved}  |  Already backed up: {already_exists}  |  Skipped: {skipped}")
    get_stats()


backup_inbox()
