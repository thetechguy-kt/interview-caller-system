import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import os
from datetime import datetime
import qrcode
import shutil

EXCEL_FILE = "candidate_list.xlsx"
TICKET_FOLDER = "Tickets"
CONFIG_FOLDER = "config"
os.makedirs(CONFIG_FOLDER, exist_ok=True)
DATE_TRACK_FILE = os.path.join(CONFIG_FOLDER, "last_ticket_date.txt")

class InterviewCandidatePOS:
    def __init__(self, root):
        self.root = root
        self.root.title("OHI IITC Candidate POS")
        self.set_window_size(420, 480)
        self.root.configure(bg="#f0f8ff")

        self.today = datetime.now().strftime("%Y-%m-%d")
        self.check_and_reset_daily()

        # Layout
        self.main_frame = tk.Frame(root, bg="#f0f8ff")
        self.main_frame.pack(fill='both', expand=True)
        self.main_frame.rowconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=0)
        self.main_frame.columnconfigure(0, weight=1)

        self.input_frame = tk.Frame(self.main_frame, bg="#f0f8ff", padx=20, pady=20)
        self.input_frame.grid(row=0, column=0, sticky='nsew')

        self.button_frame = tk.Frame(self.main_frame, bg="#f0f8ff", padx=20, pady=10)
        self.button_frame.grid(row=1, column=0, sticky='ew')

        # Configure input layout
        for i in range(6):
            self.input_frame.rowconfigure(i, weight=0)
        self.input_frame.columnconfigure(0, weight=0)
        self.input_frame.columnconfigure(1, weight=1)

        # Title Label
        tk.Label(self.input_frame, text="🎓 OHI IITC Interview Candidate POS",
                 font=("Arial", 15, "bold"), bg="#f0f8ff").grid(row=0, column=0, columnspan=2, pady=(0, 15))

        # Candidate Name
        tk.Label(self.input_frame, text="Candidate Name:", bg="#f0f8ff").grid(row=1, column=0, sticky='w', pady=5)
        self.name_entry = tk.Entry(self.input_frame)
        self.name_entry.grid(row=1, column=1, sticky='ew', pady=5)

        # Contact Number
        tk.Label(self.input_frame, text="Contact Number:", bg="#f0f8ff").grid(row=2, column=0, sticky='w', pady=5)
        self.contact_number_entry = tk.Entry(self.input_frame)
        self.contact_number_entry.grid(row=2, column=1, sticky='ew', pady=5)

        # Ticket number display
        self.ticket_label = tk.Label(self.input_frame, text=f"Candidates Logged In: {self.ticket_number}",
                                     font=("Arial", 20, "bold"), fg="#004080", bg="#f0f8ff")
        self.ticket_label.grid(row=4, column=0, columnspan=2, pady=20)

        # Buttons
        btn_generate = tk.Button(self.button_frame, text="Generate Entry Pass", font=("Arial", 12),
                                 bg="#4CAF50", fg="white", command=self.generate_ticket)
        btn_generate.pack(fill='x', pady=5)

        btn_reset = tk.Button(self.button_frame, text="Reset Counter", font=("Arial", 10),
                              bg="#f44336", fg="white", command=self.reset_counter)
        btn_reset.pack(fill='x')

        self.setup_excel()

    def set_window_size(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(400, 450)

    def check_and_reset_daily(self):
        if not os.path.exists(DATE_TRACK_FILE):
            with open(DATE_TRACK_FILE, 'w') as f:
                f.write(self.today)
            self.ticket_number = 0
        else:
            with open(DATE_TRACK_FILE, 'r') as f:
                last_date = f.read().strip()

            if last_date != self.today:
                self.ticket_number = 0
                with open(DATE_TRACK_FILE, 'w') as f:
                    f.write(self.today)
                self.reset_excel_file()
            else:
                self.ticket_number = self.get_last_ticket_number()

    def reset_excel_file(self):
        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            ws.delete_rows(2, ws.max_row)
            wb.save(EXCEL_FILE)

    def get_last_ticket_number(self):
        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            count = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] == self.today:
                    count += 1
            return count
        return 0

    def setup_excel(self):
        if not os.path.exists(EXCEL_FILE):
            wb = Workbook()
            ws = wb.active
            headers = ["Date", "Day", "Time", "Candidate Name", "Contact Number", "Entry No"]
            ws.append(headers)

            for col, header in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=col)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                ws.column_dimensions[chr(64 + col)].width = max(len(header) + 5, 15)

            wb.save(EXCEL_FILE)

    def generate_ticket(self):
        name = self.name_entry.get().strip()
        contact_number = self.contact_number_entry.get().strip()

        if not name:
            messagebox.showwarning("Input Required", "Please enter the candidate's name.")
            return

        if not contact_number:
            messagebox.showwarning("Input Required", "Please enter the contact number.")
            return

        now = datetime.now()
        date = now.strftime("%Y-%m-%d")
        day = now.strftime("%A")
        time = now.strftime("%H:%M:%S")
        file_time = now.strftime("%H-%M-%S")

        self.ticket_number += 1
        self.ticket_label.config(text=f"Entry No: {self.ticket_number}")

        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append([date, day, time, name, contact_number, self.ticket_number])
        new_row = ws.max_row
        for col in range(1, 7):
            ws.cell(row=new_row, column=col).alignment = Alignment(horizontal='center')
        wb.save(EXCEL_FILE)

        folder_name = os.path.join(TICKET_FOLDER, f"{date} - Entries")
        os.makedirs(folder_name, exist_ok=True)
        daily_excel_path = os.path.join(folder_name, f"candidate_list_{date}.xlsx")
        shutil.copy(EXCEL_FILE, daily_excel_path)

        safe_name = name.replace(" ", "_")
        pdf_filename = f"Entry_{self.ticket_number}_{safe_name}_{file_time}.pdf"
        pdf_path = os.path.join(folder_name, pdf_filename)

        self.create_ticket_pdf(pdf_path, name, contact_number, self.ticket_number, date, day, time)

        self.name_entry.delete(0, tk.END)
        self.contact_number_entry.delete(0, tk.END)

        messagebox.showinfo("Success", f"Entry No {self.ticket_number} generated for {name}.")

        if messagebox.askyesno("Print Entry Pass", "Do you want to print the pass now?"):
            try:
                os.startfile(pdf_path, "print")
            except Exception as e:
                messagebox.showerror("Printing Error", f"Could not print ticket: {e}")

    def create_ticket_pdf(self, filepath, name, contact_number, entry_no, date, day, time):
        width = 8 * cm
        height = 8 * cm
        c = canvas.Canvas(filepath, pagesize=(width, height))

        qr_text = (f"Entry No: {entry_no}\n"
                   f"Name: {name}\n"
                   f"Contact Number: {contact_number}\n"
                   f"Date: {date} ({day})\n"
                   f"Time: {time}\n")

        qr_img = qrcode.make(qr_text)
        qr_temp = "temp_qr.png"
        qr_img.save(qr_temp)

        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(width / 2, height - 30, "OHI IITC")

        c.setFont("Helvetica", 10)
        c.drawCentredString(width / 2, height - 50, f"{date} ({day}) | {time}")

        c.setFont("Helvetica", 10)
        c.drawString(20, height - 80, f"Name: {name}")
        c.drawString(20, height - 110, f"Number: {contact_number}")
        c.drawString(20, height - 140, f"Entry No: {entry_no}")

        c.drawImage(qr_temp, width - 90, 20, width=70, height=70)

        c.setFont("Helvetica-Oblique", 8)
        c.drawCentredString(width / 2, 10, "Scan for interview info")

        c.showPage()
        c.save()

        os.remove(qr_temp)

    def reset_counter(self):
        if messagebox.askyesno("Confirm Reset", "Are you sure you want to reset the entry number?"):
            self.ticket_number = 0
            self.ticket_label.config(text="Entry No: 0")
            self.reset_excel_file()

if __name__ == "__main__":
    root = tk.Tk()
    app = InterviewCandidatePOS(root)
    root.mainloop()
