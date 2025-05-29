import win32com.client
import re
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import pythoncom  # <-- DODAJ TO!

# Flaga stopowania
stop_processing = False

def process_emails(output_text, progress_bar, progress_label, count_label):
    global stop_processing
    stop_processing = False

    pythoncom.CoInitialize()

    try:
        KEYWORDS = [
            "florence", "jobrapido", "totaljobs", "daily jobs", "gowork", "uber",
            "perfectjobs4u", "ellis winters", "kasia z bolt", "codepen", "remote worker",
            "the career wallet", "internations", "arena","Paul at Contract Spy", "Allegro"
        ]

        KEYWORDS = [k.lower() for k in KEYWORDS]

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)

        target_folder_name = "NIEPRZYDATNE"
        try:
            target_folder = inbox.Folders[target_folder_name]
        except:
            target_folder = inbox.Folders.Add(target_folder_name)
            output_text.insert(tk.END, f"Utworzono folder: {target_folder_name}\n")
            output_text.update()

        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        total_messages = len(messages)
        if total_messages == 0:
            messagebox.showinfo("Informacja", "Brak wiadomości do przetworzenia.")
            return

        moved_count = 0

        for idx, message in enumerate(list(messages)):
            if stop_processing:
                output_text.insert(tk.END, "\nPrzetwarzanie zostało przerwane przez użytkownika.\n")
                break

            try:
                subject = message.Subject.lower() if message.Subject else ""
                body = message.Body.lower() if message.Body else ""
                sender_name = message.SenderName.lower() if message.SenderName else ""
                sender_email = message.SenderEmailAddress.lower() if message.SenderEmailAddress else ""

                if any(keyword in subject or keyword in body or keyword in sender_name or keyword in sender_email for keyword in KEYWORDS):
                    output_text.insert(tk.END, f"PRZENIESIONO: {message.Subject}\n")
                    output_text.update()
                    message.Move(target_folder)
                    moved_count += 1

                percent_complete = int((idx + 1) / total_messages * 100)
                progress_bar['value'] = percent_complete
                progress_label.config(text=f"{percent_complete}%")
                output_text.update()
                progress_bar.update()

            except Exception as e:
                output_text.insert(tk.END, f"Błąd przy przetwarzaniu wiadomości: {e}\n")
                output_text.update()

        output_text.insert(tk.END, f"\nSkończono! Przeniesiono {moved_count} wiadomości.\n")
        count_label.config(text=f"Przeniesione maile: {moved_count}")

    except Exception as e:
        messagebox.showerror("Błąd", f"Wystąpił błąd: {e}")


def start_processing():
    output_text.delete(1.0, tk.END)  # Wyczyść okno tekstowe
    count_label.config(text="Przeniesione maile: 0")
    progress_bar['value'] = 0
    progress_label.config(text="0%")
    threading.Thread(target=process_emails, args=(output_text, progress_bar, progress_label, count_label)).start()

def stop_processing_func():
    global stop_processing
    stop_processing = True

# Budowa GUI
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
