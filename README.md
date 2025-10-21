# Outlook Filter

A small Windows desktop utility (Python + Tkinter + Outlook COM) that scans **unread** emails in Outlook and moves messages whose **Subject** or **Sender name / email** match a **blacklist** into a folder called `NIEPRZYDATNE`.

> **Target:** Windows + Microsoft Outlook (desktop)  
> **Stack:** Python 3.10+, `pywin32` (COM), `tkinter`

---

##  Features

- Scans only **Unread** emails in the Inbox.
- Matches blacklist keywords against:
  - `Subject`
  - `SenderName`
  - `SenderEmailAddress`
  - Combined form: `SenderName <SenderEmailAddress>`
- Moves matched messages to **`NIEPRZYDATNE`** (auto-creates the folder if missing).
- Progress bar, live log and **final count** of moved emails.
- Safe iteration using Outlook `EntryID` snapshot (prevents “message has been changed” COM errors).
- Start/Stop buttons; UI remains responsive.

---

##  Requirements

- Windows 10/11
- Microsoft Outlook (desktop) configured with your account
- Python **3.10+**
- Packages: `pywin32` (installs `pythoncom`/`win32com`)

```bash
pip install pywin32
