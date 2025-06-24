import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook
import json
import os
from datetime import datetime

# Constants
EXCEL_FILE = "candidate_list.xlsx"
STATE_FILE = "queue_state.json"
COUNTER_NAME = "Room 2"  # Change for each instance

if not os.path.exists(STATE_FILE):
    with open(STATE_FILE, 'w') as f:
        json.dump({"called_tokens": []}, f)


class TokenCallerApp:
    def __init__(self, master):
        self.master = master
        self.master.title(f"{COUNTER_NAME} Control Panel")
        self.master.geometry("400x280")

        heading = tk.Label(master, text=COUNTER_NAME, font=("Arial", 18, "bold"))
        heading.pack(pady=5)

        self.token_data = []
        self.current_token = None
        self.counter_closed = False

        # Display window
        self.display_window = tk.Toplevel(master)
        self.display_window.title(f"{COUNTER_NAME} Display")
        self.display_window.geometry("300x200")
        self.display_window.protocol("WM_DELETE_WINDOW", self.on_display_close)

        self.display_heading = tk.Label(self.display_window, text="CompanyName",
                                        font=("Arial", 20, "bold"), fg="black")
        self.display_heading.pack(pady=(10, 5))

        self.counter_heading = tk.Label(self.display_window, text=COUNTER_NAME,
                                        font=("Arial", 14, "bold"), fg="black")
        self.counter_heading.pack(pady=(0, 10))

        self.display_label = tk.Label(self.display_window, text="Waiting...",
                                      font=("Arial", 24, "bold"), fg="blue")
        self.display_label.pack(expand=True)

        # Control buttons
        self.call_button = tk.Button(master, text="Call Next", font=("Arial", 14),
                                     command=self.call_next)
        self.call_button.pack(pady=5, fill='x')

        self.recall_button = tk.Button(master, text="Recall", font=("Arial", 12),
                                       command=self.recall)
        self.recall_button.pack(pady=5, fill='x')

        self.waiting_button = tk.Button(master, text="Waiting", font=("Arial", 12),
                                        command=self.set_waiting)
        self.waiting_button.pack(pady=5, fill='x')

        self.close_button = tk.Button(master, text="Close Room", font=("Arial", 12),
                                      fg="white", bg="red", command=self.close_counter)
        self.close_button.pack(pady=5, fill='x')

        self.open_button = tk.Button(master, text="Open Room", font=("Arial", 12),
                                     fg="white", bg="green", command=self.open_counter)
        self.open_button.pack(pady=5, fill='x')

        self.token_label = tk.Label(master, text="Token: -\nName: -", font=("Arial", 12))
        self.token_label.pack(pady=5)

        self.load_tokens()
        self.refresh_excel_data()

    def on_display_close(self):
        messagebox.showinfo("Info", "Display window cannot be closed separately.")

    def load_tokens(self):
        self.token_data = []
        if not os.path.exists(EXCEL_FILE):
            messagebox.showerror("Missing File", f"{EXCEL_FILE} not found.")
            return

        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        today = datetime.now().strftime("%Y-%m-%d")

        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == today:
                self.token_data.append({
                    "token": row[5],  # Corrected from row[7] to row[5]
                    "name": row[3],
                    "date": row[0],
                    "time": row[2]
                })

    def refresh_excel_data(self):
        if not self.counter_closed:
            self.load_tokens()
        self.master.after(3000, self.refresh_excel_data)

    def call_next(self):
        if self.counter_closed:
            messagebox.showwarning("Room Closed", "This room is closed.")
            return

        with open(STATE_FILE, 'r') as f:
            state = json.load(f)

        called_tokens = [item['token'] for item in state.get("called_tokens", [])]

        next_token = next((t for t in self.token_data if t["token"] not in called_tokens), None)

        if next_token:
            self.current_token = next_token
            state["called_tokens"].append({
                "token": next_token["token"],
                "name": next_token["name"],
                "counter": COUNTER_NAME,
                "time": next_token["time"]
            })
            with open(STATE_FILE, 'w') as f:
                json.dump(state, f, indent=2)
            self.update_display(next_token)
        else:
            messagebox.showinfo("Info", "No more tokens to call.")

    def recall(self):
        if self.counter_closed:
            messagebox.showwarning("Room Closed", "This room is closed.")
            return

        if self.current_token:
            self.update_display(self.current_token)
        else:
            messagebox.showinfo("Info", "No token to recall.")

    def set_waiting(self):
        if self.counter_closed:
            messagebox.showwarning("Room Closed", "This room is closed.")
            return

        self.current_token = None
        self.token_label.config(text="Waiting", fg="blue")
        self.display_label.config(text="Waiting", fg="blue", font=("Arial", 24, "bold"))

    def update_display(self, token_info):
        token_text = f"Token: {token_info['token']}\nName: {token_info['name']}"
        self.token_label.config(text=token_text, fg="black")
        self.display_label.config(text=f"Token {token_info['token']}\n{token_info['name']}",
                                  fg="blue", font=("Arial", 24, "bold"))

    def close_counter(self):
        if messagebox.askyesno("Close Room", "Are you sure you want to close this interview room?"):
            self.counter_closed = True
            self.call_button.config(state='disabled')
            self.recall_button.config(state='disabled')
            self.waiting_button.config(state='disabled')
            self.close_button.config(state='disabled')

            self.display_label.config(text="Room Closed", fg="red", font=("Arial", 28, "bold"))
            self.token_label.config(text="Room Closed", fg="red")

    def open_counter(self):
        if not self.counter_closed:
            messagebox.showinfo("Info", "Room is already open.")
            return

        self.counter_closed = False
        self.call_button.config(state='normal')
        self.recall_button.config(state='normal')
        self.waiting_button.config(state='normal')
        self.close_button.config(state='normal')

        self.current_token = None
        self.token_label.config(text="Waiting", fg="blue")
        self.display_label.config(text="Waiting", fg="blue", font=("Arial", 24, "bold"))


if __name__ == "__main__":
    root = tk.Tk()
    app = TokenCallerApp(root)
    root.mainloop()
