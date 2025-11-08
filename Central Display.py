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

# Colors for dark theme with contrast
BG_COLOR = "#1e1e1e"        # Dark gray background
FG_COLOR = "#e0e0e0"        # Light gray text
HEADER_BG_COLOR = "#004080" # Deep blue for header
SELECT_BG_COLOR = "#3399ff" # Soft blue for selection and blinking
ROW_COLOR_1 = "#1e1e1e"     # Odd row background
ROW_COLOR_2 = "#2a2a2a"     # Even row background
SELECT_FG_COLOR = "#ffffff" # White text for selected/highlighted


class CentralDisplayApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KTech Central Display")
        self.root.configure(bg=BG_COLOR)

        self.root.attributes('-fullscreen', True)
        self.root.bind("<Escape>", self.exit_fullscreen)
        self.root.bind("<f>", self.enter_fullscreen)
        self.root.bind("<F>", self.enter_fullscreen)

        # Header Label
        tk.Label(root, text="🎓 KTech Interview",
                 font=("Arial", 36, "bold"),
                 fg=FG_COLOR,
                 bg=BG_COLOR).pack(pady=(20, 5))

        # Fullscreen button
        btn_fullscreen = tk.Button(
            root, text="🔲 Fullscreen (F)", font=("Arial", 10),
            command=self.enter_fullscreen,
            bg=HEADER_BG_COLOR, fg="white",
            activebackground="#0059b3", activeforeground="white"
        )
        btn_fullscreen.pack(anchor="ne", padx=20, pady=(5, 0))

        # Time label (bolder, bright)
        self.time_label = tk.Label(root, text="", font=("Arial", 22, "bold"),
                                   bg=BG_COLOR, fg=FG_COLOR)
        self.time_label.pack()

        # Treeview columns
        columns = ("Token", "Name", "Room")
        self.tree = ttk.Treeview(root, columns=columns, show="headings", height=15)
        self.tree.heading("Token", text="Token No")
        self.tree.heading("Name", text="Candidate Name")
        self.tree.heading("Room", text="Room")

        self.tree.column("Token", anchor="center", width=300)
        self.tree.column("Name", anchor="center", width=500)
        self.tree.column("Room", anchor="center", width=300)

        self.tree.pack(pady=40, expand=True, fill='both')

        # Style the treeview
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Treeview",
                        background=BG_COLOR,
                        foreground=FG_COLOR,
                        fieldbackground=BG_COLOR,
                        font=("Arial", 18),
                        rowheight=50)
        style.configure("Treeview.Heading",
                        font=("Arial", 26, "bold"),
                        background=HEADER_BG_COLOR,
                        foreground="white")
        style.map('Treeview', background=[('selected', SELECT_BG_COLOR)],
                  foreground=[('selected', SELECT_FG_COLOR)])

        # Row tags for alternating colors
        self.tree.tag_configure('oddrow', background=ROW_COLOR_1, foreground=FG_COLOR)
        self.tree.tag_configure('evenrow', background=ROW_COLOR_2, foreground=FG_COLOR)
        self.tree.tag_configure('blink', background=SELECT_BG_COLOR, foreground=SELECT_FG_COLOR)

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
                key=lambda x: x[1].get("timestamp", ""),
                reverse=True
            )

            for i, (counter, entry) in enumerate(sorted_items):
                token = entry["token"]
                name = entry["name"]
                latest_data[counter] = token

                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                row_id = self.tree.insert("", "end", values=(token, name, counter), tags=(tag,))

                # If token changed, play sound + blink
                if counter not in self.previous_data or self.previous_data[counter] != token:
                    self.blink_row(row_id, 0)
                    self.play_sound()

        self.previous_data = latest_data
        self.root.after(REFRESH_INTERVAL, self.refresh_data)

    def blink_row(self, row_id, count):
        if count < 6:
            tags = self.tree.item(row_id, "tags")

            normal_bg = None
            for tag in tags:
                if tag == 'oddrow':
                    normal_bg = ROW_COLOR_1
                    break
                elif tag == 'evenrow':
                    normal_bg = ROW_COLOR_2
                    break
            if normal_bg is None:
                normal_bg = BG_COLOR

            # Toggle color between highlight and normal
            color = SELECT_BG_COLOR if count % 2 == 0 else normal_bg
            self.tree.tag_configure("blink", background=color, foreground=SELECT_FG_COLOR)

            # Replace odd/even row tag with blink tag
            new_tags = [t for t in tags if t not in ('oddrow', 'evenrow')]
            new_tags.append("blink")
            self.tree.item(row_id, tags=new_tags)

            self.root.after(500, lambda: self.blink_row(row_id, count + 1))
        else:
            # Restore original tags
            tags = self.tree.item(row_id, "tags")
            new_tags = [t for t in tags if t != "blink"]
            index = self.tree.index(row_id)
            original_tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            if original_tag not in new_tags:
                new_tags.append(original_tag)
            self.tree.item(row_id, tags=new_tags)

    def play_sound(self):
        if platform.system() == "Windows" and os.path.exists(SOUND_FILE):
            winsound.PlaySound(SOUND_FILE, winsound.SND_FILENAME | winsound.SND_ASYNC)


if __name__ == "__main__":
    root = tk.Tk()
    app = CentralDisplayApp(root)
    root.mainloop()
