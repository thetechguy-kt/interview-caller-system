import tkinter as tk
from tkinter import ttk
import json
import os
from datetime import datetime
import platform

# Only for Windows sound
if platform.system() == "Windows":
    import winsound

STATE_FILE = "queue_state.json"
SOUND_FILE = "dip_config/notify.wav"  # Ensure this file exists
REFRESH_INTERVAL = 3000  # milliseconds

class CentralDisplayApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CompanyName Central Display")
        self.root.configure(bg="#f0f8ff")

        self.root.attributes('-fullscreen', True)
        self.root.bind("<Escape>", self.exit_fullscreen)
        self.root.bind("<f>", self.enter_fullscreen)
        self.root.bind("<F>", self.enter_fullscreen)

        tk.Label(root, text="🎓 CompanyName Interview", font=("Arial", 36, "bold"),
                 fg="#004080", bg="#f0f8ff").pack(pady=(20, 5))

        btn_fullscreen = tk.Button(
            root, text="🔲 Fullscreen (F)", font=("Arial", 10),
            command=self.enter_fullscreen, bg="#004080", fg="white"
        )
        btn_fullscreen.pack(anchor="ne", padx=20, pady=(5, 0))

        self.time_label = tk.Label(root, text="", font=("Arial", 18),
                                   bg="#f0f8ff", fg="#333")
        self.time_label.pack()

        columns = ("Token", "Name", "Room")
        self.tree = ttk.Treeview(root, columns=columns, show="headings", height=15)
        self.tree.heading("Token", text="Token No")
        self.tree.heading("Name", text="Candidate Name")
        self.tree.heading("Room", text="Room")

        self.tree.column("Token", anchor="center", width=300)
        self.tree.column("Name", anchor="center", width=500)
        self.tree.column("Room", anchor="center", width=300)

        self.tree.pack(pady=40, expand=True)

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 26, "bold"))
        style.configure("Treeview", font=("Arial", 24), rowheight=60)

        self.previous_data = {}
        self.update_time()
        self.refresh_data()

    def exit_fullscreen(self, event=None):
        self.root.attributes('-fullscreen', False)

    def enter_fullscreen(self, event=None):
        self.root.attributes('-fullscreen', True)

    def update_time(self):
        now = datetime.now().strftime("%A, %d %B %Y  |  %I:%M:%S %p")
        self.time_label.config(text=now)
        self.root.after(1000, self.update_time)

    def refresh_data(self):
        self.tree.delete(*self.tree.get_children())
        latest_data = {}

        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, "r") as f:
                state = json.load(f)

            latest_per_counter = {}
            for item in state.get("called_tokens", []):
                latest_per_counter[item["counter"]] = item

            # Sort to show most recent token first
            sorted_items = sorted(
                latest_per_counter.items(),
                key=lambda x: x[1].get("timestamp", ""), reverse=True
            )

            for counter, entry in sorted_items:
                token = entry["token"]
                name = entry["name"]
                latest_data[counter] = token

                row_id = self.tree.insert("", 0, values=(token, name, counter))

                # If token changed, play sound + blink
                if counter not in self.previous_data or self.previous_data[counter] != token:
                    self.blink_row(row_id, 0)
                    self.play_sound()

        self.previous_data = latest_data
        self.root.after(REFRESH_INTERVAL, self.refresh_data)

    def blink_row(self, row_id, count):
        if count < 6:
            color = "#ffeb3b" if count % 2 == 0 else "white"
            self.tree.item(row_id, tags=("blink",))
            self.tree.tag_configure("blink", background=color)
            self.root.after(500, lambda: self.blink_row(row_id, count + 1))
        else:
            self.tree.tag_configure("blink", background="white")

    def play_sound(self):
        if platform.system() == "Windows" and os.path.exists(SOUND_FILE):
            winsound.PlaySound(SOUND_FILE, winsound.SND_FILENAME | winsound.SND_ASYNC)

if __name__ == "__main__":
    root = tk.Tk()
    app = CentralDisplayApp(root)
    root.mainloop()
