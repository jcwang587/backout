import subprocess
import sys
import os
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils import BASE_DIR, init_db, save_email, get_stats


def run(cmd, **kwargs):
    return subprocess.run(cmd, capture_output=True, text=True, **kwargs)


def quit_outlook():
    run(["osascript", "-e", 'tell application "Microsoft Outlook" to quit'])
    time.sleep(3)


def set_outlook_mode(new_outlook: bool):
    value = "1" if new_outlook else "0"
    run(["defaults", "write", "com.microsoft.Outlook", "IsRunningNewOutlook", "-bool", value])


def open_outlook():
    run(["open", "-a", "Microsoft Outlook"])


def wait_for_outlook_ready(timeout=90):
    print("Waiting for Legacy Outlook to load", end="", flush=True)
    start = time.time()
    while time.time() - start < timeout:
        result = run(["osascript", "-e", """
            tell application "Microsoft Outlook"
                set allFolders to every mail folder
                repeat with f in allFolders
                    if (name of f is "Inbox") and ((count of messages of f) > 0) then
                        return count of messages of f
                    end if
                end repeat
                return 0
            end tell
        """])
        if result.returncode == 0 and result.stdout.strip().isdigit() and int(result.stdout.strip()) > 0:
            print(f" ready ({result.stdout.strip()} emails).")
            return True
        print(".", end="", flush=True)
        time.sleep(5)
    print(" timed out.")
    return False


def backup_inbox():
    script = """
tell application "Microsoft Outlook"
    set output to ""
    set allFolders to every mail folder
    set inboxFolder to missing value
    repeat with f in allFolders
        if (name of f is "Inbox") and ((count of messages of f) > 0) then
            set inboxFolder to f
            exit repeat
        end if
    end repeat
    if inboxFolder is missing value then return "ERROR: inbox not found"
    set msgs to messages of inboxFolder
    repeat with m in msgs
        set msgDate to (time sent of m) as string
        set msgSubject to subject of m
        set msgBody to plain text content of m
        set msgSender to sender of m
        set sAddr to address of msgSender
        set sName to sAddr
        try
            set sName to name of msgSender
        end try
        set output to output & "%%DATE%%:" & msgDate & "\n"
        set output to output & "%%FROM%%:" & sName & " <" & sAddr & ">\n"
        set output to output & "%%SUBJECT%%:" & msgSubject & "\n"
        set output to output & "%%BODY%%:" & msgBody & "\n"
        set output to output & "%%END%%\n"
    end repeat
    return output
end tell
"""
    result = run(["osascript", "-e", script])
    if result.returncode != 0:
        print("Backup error:", result.stderr)
        return

    con = init_db()
    saved = skipped = already_exists = 0
    for block in result.stdout.split("%%END%%\n"):
        block = block.strip()
        if not block:
            continue
        try:
            date_str = block.split("%%DATE%%:")[1].split("\n%%FROM%%:")[0]
            sender   = block.split("%%FROM%%:")[1].split("\n%%SUBJECT%%:")[0]
            subject  = block.split("%%SUBJECT%%:")[1].split("\n%%BODY%%:")[0]
            body     = block.split("%%BODY%%:")[1]
            if save_email(date_str, sender, subject, body, con):
                saved += 1
            else:
                already_exists += 1
        except Exception:
            skipped += 1
    con.close()

    print(f"New: {saved}  |  Already backed up: {already_exists}  |  Skipped: {skipped}")


print("=== Outlook Email Backup (Mac) ===")

print("Step 1: Switching to Legacy Outlook...")
quit_outlook()
set_outlook_mode(new_outlook=False)
open_outlook()

if not wait_for_outlook_ready():
    print("Could not connect to Outlook. Restoring New Outlook and exiting.")
    quit_outlook()
    set_outlook_mode(new_outlook=True)
    open_outlook()
    exit(1)

print("Step 2: Backing up emails...")
backup_inbox()

print("Step 3: Switching back to New Outlook...")
quit_outlook()
set_outlook_mode(new_outlook=True)
open_outlook()

print("Done. New Outlook is now open.")
get_stats()
