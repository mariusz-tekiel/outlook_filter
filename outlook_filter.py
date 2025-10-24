import win32com.client
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import pythoncom

# Flaga stopowania
stop_processing = False

# --- Aktualizacje GUI z wątku roboczego ---
def append_text_safe(widget, text):
    widget.after(0, lambda: (widget.insert(tk.END, text), widget.see(tk.END)))

def set_label_text_safe(label, text):
    label.after(0, lambda: label.config(text=text))

def set_progress_safe(bar, value):
    bar.after(0, lambda: bar.config(value=value))

def normalize(s: str) -> str:
    return (s or "").strip().lower()

def process_emails(output_text, progress_bar, progress_label, count_label):
    global stop_processing
    stop_processing = False

    pythoncom.CoInitialize()
    try:
        # --- CZARNA LISTA ---
        KEYWORDS = [
            "florence", "jobrapido", "totaljobs", "daily jobs", "gowork", "uber",
            "perfectjobs4u", "ellis winters", "kasia z bolt", "codepen", "remote worker",
            "the career wallet", "internations", "arena", "paul at contract spy", "allegro",
            "bhanu ahluwalia","Temu","Pracuj.pl","OLX","Jooble", "TeePublic", "Outside Spy",
            "Jobs from 4 Steps", "Kaufland Card","Laura at Rightmove",
            "Campaign", "Freenow","Ian Bremmer","Yanosik", "You have", "LinkedIn <notifications-noreply@linkedin.com>",
            "LinkedIn <messages-noreply@linkedin.com>","Job Placements Jobs <info@jobplacements.com>",
            "Richard Branson via LinkedIn","Rightmove Partners","LinkedIn Job Alerts <jobalerts-noreply@linkedin.com>","Redbubble",
            "Bounce","LinkedIn <updates-noreply@linkedin.com>", "cyberFolks"
            # możesz wrzucać pełne formy typu:
            # "job placements jobs <info@jobplacements.com>",
        ]
        KEYWORDS = [k.lower() for k in KEYWORDS]

        # Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox

        # Folder docelowy
        target_folder_name = "NIEPRZYDATNE"
        try:
            target_folder = inbox.Folders[target_folder_name]
        except Exception:
            target_folder = inbox.Folders.Add(target_folder_name)
            append_text_safe(output_text, f"Utworzono folder: {target_folder_name}\n")

        # --- Migawka ID nieprzeczytanych, żeby kolekcja nie rozjechała pętli ---
        items = inbox.Items.Restrict("[Unread] = True")
        items.Sort("[ReceivedTime]", True)
        total = items.Count

        if total == 0:
            messagebox.showinfo("Informacja", "Brak NIEPRZECZYTANYCH wiadomości do przetworzenia.")
            return

        # Zapisz EntryID wszystkich kandydatów, a dopiero potem je pobieraj pojedynczo
        entry_ids = []
        for i in range(1, total + 1):
            try:
                entry_ids.append(items.Item(i).EntryID)
            except Exception:
                # jeśli któryś element zniknie/zmieni się w trakcie snapshotu – pomiń
                pass

        moved_count = 0
        processed = 0
        total_ids = len(entry_ids)

        for entry_id in entry_ids:
            if stop_processing:
                append_text_safe(output_text, "\nPrzetwarzanie zostało przerwane przez użytkownika.\n")
                break

            try:
                msg = outlook.GetItemFromID(entry_id)
                if msg is None:
                    processed += 1
                    continue

                subject = normalize(getattr(msg, "Subject", ""))
                sender_name = normalize(getattr(msg, "SenderName", ""))
                sender_email = normalize(getattr(msg, "SenderEmailAddress", ""))

                # Złożona postać "Name <email>" – żeby działały pełne wpisy z czarnej listy
                name_email_combo = f"{sender_name} <{sender_email}>".lower()

                haystacks = (subject, sender_name, sender_email, name_email_combo)

                if any(any(k in h for h in haystacks) for k in KEYWORDS):
                    append_text_safe(output_text, f"PRZENIESIONO: {msg.Subject} (nadawca: {msg.SenderName})\n")
                    msg.Move(target_folder)
                    moved_count += 1

            except Exception as e:
                # Typowy COM error przy zmianie elementu przez Outlooka/serwer w trakcie – po prostu logujemy i lecimy dalej
                append_text_safe(output_text, f"Błąd przy przetwarzaniu wiadomości: {e}\n")

            processed += 1
            percent_complete = int(processed / max(1, total_ids) * 100)
            set_progress_safe(progress_bar, percent_complete)
            set_label_text_safe(progress_label, f"{percent_complete}%")

        append_text_safe(output_text, f"\nSkończono! Przeniesiono {moved_count} wiadomości.\n")
        set_label_text_safe(count_label, f"Przeniesione maile: {moved_count}")

    except Exception as e:
        messagebox.showerror("Błąd", f"Wystąpił błąd: {e}")
    finally:
        pythoncom.CoUninitialize()

def start_processing():
    output_text.delete(1.0, tk.END)
    count_label.config(text="Przeniesione maile: 0")
    progress_bar['value'] = 0
    progress_label.config(text="0%")
    threading.Thread(target=process_emails, args=(output_text, progress_bar, progress_label, count_label), daemon=True).start()

def stop_processing_func():
    global stop_processing
    stop_processing = True

# --- GUI ---
window = tk.Tk()
window.title("Outlook Cleaner - Przenoszenie NIEPRZYDATNYCH maili")

frame_buttons = tk.Frame(window)
frame_buttons.pack(pady=10)

start_button = tk.Button(frame_buttons, text="Start", command=start_processing)
start_button.grid(row=0, column=0, padx=5)

stop_button = tk.Button(frame_buttons, text="Stop", command=stop_processing_func)
stop_button.grid(row=0, column=1, padx=5)

progress_bar = ttk.Progressbar(window, length=400, mode='determinate')
progress_bar.pack(pady=5)

progress_label = tk.Label(window, text="0%")
progress_label.pack(pady=5)

count_label = tk.Label(window, text="Przeniesione maile: 0")
count_label.pack(pady=5)

output_text = scrolledtext.ScrolledText(window, width=100, height=30)
output_text.pack(padx=10, pady=10)

window.mainloop()
