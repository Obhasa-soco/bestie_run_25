import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
from typing import Iterator
from transport import SerialTransport
from reader import Reader
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from datetime import datetime
import os
from dotenv import load_dotenv
import json

# Load environment variables
load_dotenv()

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_KEY")
SHEET_ID = os.getenv("SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME")
JSON_FILE = 'attendance.json'

def load_existing_tags():
    try:
        with open(JSON_FILE, 'r') as file:
            return set(json.load(file))
    except FileNotFoundError:
        return set()

def save_tags(tags):
    existing_tags = load_existing_tags()
    new_tags = tags - existing_tags
    if new_tags:
        updated_tags = existing_tags.union(new_tags)
        with open(JSON_FILE, 'w') as file:
            json.dump(list(updated_tags), file, indent=4)

# Google Sheets setup
credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

def get_google_sheet_data():
    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=f"{SHEET_NAME}!A:E"
    ).execute()
    return result.get('values', [])

def update_attendance_sheet(detected_uids):
    rows = get_google_sheet_data()
    updates = []
    for i, row in enumerate(rows):
        if row[0] in detected_uids and row[4] == "NO":
            current_time = datetime.now().strftime("%H:%M:%S")
            updates.append({
                "range": f"{SHEET_NAME}!D{i + 1}:E{i + 1}",
                "values": [[current_time, "YES"]]
            })

    if updates:
        body = {"valueInputOption": "RAW", "data": updates}
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_ID, body=body
        ).execute()
    return len(updates) > 0

def reader_thread(port, baudrate, detected_uids, lock, update_uid_count, update_rfid_display, running_flag):
    transport = SerialTransport(port, baudrate)
    reader = Reader(transport)

    try:
        while running_flag[0]:
            tags: Iterator[bytes] = reader.inventory_answer_mode()
            for tag in tags:
                if not running_flag[0]:
                    break
                hex_tag = ''.join([f'{byte:02X}' for byte in tag])
                with lock:
                    detected_uids.add(hex_tag)
                update_uid_count(len(detected_uids))
                update_rfid_display(hex_tag)

                # Google Sheets Update (batch every 5 seconds)
                if time.time() - app.last_update_time > 5:
                    with lock:
                        if update_attendance_sheet(detected_uids):
                            app.total_updates += len(detected_uids)
                            detected_uids.clear()  # Avoid redundant updates
                        app.last_update_time = time.time()
                time.sleep(0.05)
    except Exception as e:
        print(f"Error in reader_thread: {e}")
    finally:
        reader.close()


class RFIDApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bestie Run RFID System")

        self.root.geometry("400x400")
        self.root.iconbitmap(os.getenv("ICON_BIT_MAP"))

        self.detected_uids = set()
        self.lock = threading.Lock()
        self.last_update_time = time.time()
        self.total_updates = 0

        self.reader_configs = []
        self.threads = []
        self.running = [False]  # Shared flag to control threads

        self.create_widgets()

    def create_widgets(self):
        frame = tk.Frame(self.root)
        frame.pack(pady=10)

        tk.Label(frame, text="Port 1:").grid(row=0, column=0, padx=5, pady=5)
        self.port1 = ttk.Entry(frame)
        self.port1.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame, text="Port 2:").grid(row=1, column=0, padx=5, pady=5)
        self.port2 = ttk.Entry(frame)
        self.port2.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame, text="Port 3:").grid(row=2, column=0, padx=5, pady=5)
        self.port3 = ttk.Entry(frame)
        self.port3.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(frame, text="Port 4:").grid(row=3, column=0, padx=5, pady=5)
        self.port4 = ttk.Entry(frame)
        self.port4.grid(row=3, column=1, padx=5, pady=5)

        self.start_button = ttk.Button(self.root, text="Start", command=self.start_readers)
        self.start_button.pack(pady=10)

        self.stop_button = ttk.Button(self.root, text="Stop", command=self.stop_readers, state="disabled")
        self.stop_button.pack(pady=10)

        self.uid_count_label = tk.Label(self.root, text="Total UIDs Updated: 0", font=("Arial", 14))
        self.uid_count_label.pack(pady=10)

        self.rfid_listbox = tk.Listbox(self.root, height=10, width=50)
        self.rfid_listbox.pack(pady=10)

    def update_uid_count(self, count):
        self.uid_count_label.config(text=f"Total UIDs Updated: {self.total_updates}")

    def update_rfid_display(self, tag):
        if self.running[0]:
            self.rfid_listbox.insert(tk.END, tag)
            self.rfid_listbox.see(tk.END)

    def start_readers(self):
        if self.running[0]:
            messagebox.showwarning("Warning", "Readers are already running.")
            return

        self.reader_configs = []
        for port_entry in [self.port1, self.port2, self.port3, self.port4]:
            port = port_entry.get().strip()
            if port:
                self.reader_configs.append({"port": port, "baudrate": 57600})

        if not self.reader_configs:
            messagebox.showerror("Error", "Please enter at least one port.")
            return

        self.running[0] = True
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")

        self.threads = []
        for config in self.reader_configs:
            thread = threading.Thread(target=reader_thread, args=(
                config["port"], config["baudrate"], self.detected_uids, self.lock, self.update_uid_count, self.update_rfid_display, self.running))
            thread.daemon = True
            thread.start()
            self.threads.append(thread)

    def stop_readers(self):
        if not self.running[0]:
            messagebox.showwarning("Warning", "Readers are not running.")
            return

        self.running[0] = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")

        for thread in self.threads:
            thread.join(timeout=1)
        self.threads = []

        messagebox.showinfo("Info", "Readers stopped.")

if __name__ == "__main__":
    root = tk.Tk()
    app = RFIDApp(root)
    root.mainloop()
