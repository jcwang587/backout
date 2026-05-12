import subprocess
import sys
import os
import time
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils import init_db, save_email, get_stats


def run(cmd, **kwargs):
    try:
        return subprocess.run(cmd, capture_output=True, text=True, **kwargs)
    except subprocess.TimeoutExpired as exc:
        return subprocess.CompletedProcess(
            cmd,
            124,
            stdout=exc.stdout or "",
            stderr=exc.stderr or "Command timed out",
        )


def quit_outlook():
    run(["osascript", "-e", 'tell application "Microsoft Outlook" to quit'])
    time.sleep(3)


def set_outlook_mode(new_outlook: bool):
    # EnableNewOutlook is the preference Outlook actually uses to choose the UI.
    enable_value = "2" if new_outlook else "0"
    running_value = "true" if new_outlook else "false"
    commands = [
        ["defaults", "write", "com.microsoft.Outlook", "EnableNewOutlook", "-int", enable_value],
        ["defaults", "write", "com.microsoft.Outlook", "IsRunningNewOutlook", "-bool", running_value],
        ["defaults", "write", "com.microsoft.Outlook", "RunningNewOutlook", "-bool", running_value],
    ]
    for cmd in commands:
        result = run(cmd)
        if result.returncode != 0:
            raise RuntimeError(f"Could not switch Outlook mode: {result.stderr.strip()}")


def open_outlook():
    run(["open", "-a", "Microsoft Outlook"])


def inbox_snapshot():
    script = """
tell application "Microsoft Outlook"
    set totalCount to 0
    set inboxCount to 0
    set currentMonthCount to 0
    set newestDate to missing value
    set nowDate to current date
    set allFolders to every mail folder
    repeat with f in allFolders
        if (name of f is "Inbox") then
            set inboxCount to inboxCount + 1
            set folderMessages to messages of f
            set totalCount to totalCount + (count of folderMessages)
            repeat with m in folderMessages
                set msgDate to missing value
                try
                    set msgDate to time received of m
                on error
                    try
                        set msgDate to time sent of m
                    end try
                end try
                if msgDate is not missing value then
                    if newestDate is missing value or msgDate > newestDate then
                        set newestDate to msgDate
                    end if
                    if ((year of msgDate) is (year of nowDate)) and ((month of msgDate) is (month of nowDate)) then
                        set currentMonthCount to currentMonthCount + 1
                    end if
                end if
            end repeat
        end if
    end repeat
    set newestText to "none"
    if newestDate is not missing value then set newestText to newestDate as string
    return (totalCount as string) & "|" & (inboxCount as string) & "|" & (currentMonthCount as string) & "|" & newestText
end tell
"""
    result = run(["osascript", "-e", script])
    if result.returncode != 0:
        return {"total": 0, "inboxes": 0, "current_month": 0, "newest": "none"}
    try:
        total, inboxes, current_month, newest = result.stdout.strip().split("|", 3)
        return {
            "total": int(total),
            "inboxes": int(inboxes),
            "current_month": int(current_month),
            "newest": newest,
        }
    except ValueError:
        return {"total": 0, "inboxes": 0, "current_month": 0, "newest": "none"}


def wait_for_outlook_ready(timeout=90):
    print("Waiting for Legacy Outlook to load", end="", flush=True)
    start = time.time()
    while time.time() - start < timeout:
        snapshot = inbox_snapshot()
        if snapshot["total"] > 0:
            print(f" ready ({snapshot['total']} emails).")
            return True
        print(".", end="", flush=True)
        time.sleep(5)
    print(" timed out.")
    return False


def request_outlook_sync():
    result = run(["osascript", "-e", 'tell application "Microsoft Outlook" to sync'], timeout=30)
    return result.returncode == 0


def sync_outlook(timeout=120, settle_time=15):
    print("Syncing emails", end="", flush=True)
    if not request_outlook_sync():
        print(" could not request sync; continuing with currently visible emails.")
        return inbox_snapshot()

    start = time.time()
    last_snapshot = inbox_snapshot()
    last_change = time.time()
    while time.time() - start < timeout:
        print(".", end="", flush=True)
        time.sleep(5)
        snapshot = inbox_snapshot()
        if snapshot != last_snapshot:
            last_snapshot = snapshot
            last_change = time.time()
        elif time.time() - last_change >= settle_time:
            break

    print(f" done ({last_snapshot['total']} emails).")
    return last_snapshot


def backup_inbox():
    script = """
tell application "Microsoft Outlook"
    set output to ""
    set allFolders to every mail folder
    repeat with f in allFolders
        if (name of f is "Inbox") and ((count of messages of f) > 0) then
            set msgs to messages of f
            repeat with m in msgs
                set msgDate to (time sent of m) as string

                try
                    set msgSubject to subject of m
                on error
                    set msgSubject to "(no subject)"
                end try
                if msgSubject is missing value then set msgSubject to "(no subject)"

                try
                    set msgBody to plain text content of m
                on error
                    set msgBody to ""
                end try
                if msgBody is missing value then set msgBody to ""

                set sAddr to ""
                set sName to "(unknown sender)"
                try
                    set msgSender to sender of m
                    set sAddr to address of msgSender
                    set sName to sAddr
                    try
                        set sName to name of msgSender
                    end try
                end try

                set output to output & "%%DATE%%:" & msgDate & "\n"
                set output to output & "%%FROM%%:" & sName & " <" & sAddr & ">\n"
                set output to output & "%%SUBJECT%%:" & msgSubject & "\n"
                set output to output & "%%BODY%%:" & msgBody & "\n"
                set output to output & "%%END%%\n"
            end repeat
        end if
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

    snapshot = inbox_snapshot()
    if snapshot["current_month"] == 0:
        current_month = datetime.now().strftime("%Y-%m")
        print(
            f"Warning: Legacy Outlook did not expose any {current_month} Inbox messages "
            f"(newest visible: {snapshot['newest']}). "
            "If you see May mail in New Outlook, manually refresh/sync Legacy Outlook, then run this again."
        )


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

print("Step 2: Syncing and backing up emails...")
sync_outlook()
backup_inbox()

print("Step 3: Switching back to New Outlook...")
quit_outlook()
set_outlook_mode(new_outlook=True)
open_outlook()

print("Done. New Outlook is now open.")
get_stats()
