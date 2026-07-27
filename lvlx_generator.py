import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import sqlite3
import csv
import shutil
import re
import zipfile
import random
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm

# --- MAC .APP DIRECTORY FIX ---
if getattr(sys, 'frozen', False):
    if sys.platform == 'darwin' and '.app' in sys.executable:
        app_dir = os.path.dirname(os.path.dirname(os.path.dirname(sys.executable)))
        os.chdir(os.path.dirname(app_dir))
    else:
        os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

# --- PDF & Image Conversion Libraries ---
try:
    from docx2pdf import convert as convert_to_pdf
    HAS_DOCX2PDF = True
except ImportError:
    HAS_DOCX2PDF = False

try:
    from pypdf import PdfWriter, PdfReader
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

try:
    from PIL import Image
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False

# Gender-specific WHO Medians
WHO_GUIDELINES = {
    'M': {
        'BMI': {5: 15.2, 6: 15.3, 7: 15.5, 8: 15.7, 9: 16.1, 10: 16.7, 11: 17.3, 12: 18.0, 13: 18.8, 14: 19.6, 15: 20.4, 16: 21.1, 17: 21.8, 18: 22.4},
        'Height': {5: 110.0, 6: 116.0, 7: 121.9, 8: 127.3, 9: 132.6, 10: 137.8, 11: 143.1, 12: 149.1, 13: 156.0, 14: 163.2, 15: 170.0, 16: 173.5, 17: 175.2, 18: 176.1},
        'Weight': {5: 18.3, 6: 20.5, 7: 22.9, 8: 25.4, 9: 28.1, 10: 31.2, 11: 35.3, 12: 39.8, 13: 45.3, 14: 50.8, 15: 56.0, 16: 60.8, 17: 64.6, 18: 66.9}
    },
    'F': {
        'BMI': {5: 15.3, 6: 15.3, 7: 15.4, 8: 15.7, 9: 16.1, 10: 16.6, 11: 17.2, 12: 18.0, 13: 18.8, 14: 19.6, 15: 20.2, 16: 20.7, 17: 21.1, 18: 21.4},
        'Height': {5: 109.4, 6: 115.1, 7: 120.8, 8: 126.6, 9: 132.2, 10: 138.6, 11: 145.0, 12: 151.2, 13: 156.4, 14: 159.8, 15: 161.7, 16: 162.5, 17: 162.9, 18: 163.1},
        'Weight': {5: 18.2, 6: 20.2, 7: 22.4, 8: 25.0, 9: 28.2, 10: 31.9, 11: 36.9, 12: 41.5, 13: 45.8, 14: 49.8, 15: 53.0, 16: 55.4, 17: 56.7, 18: 57.3}
    }
}

def get_healthy_range(metric, age, gender):
    if gender not in ['M', 'F'] or age is None: return "N/A"
    safe_age = max(5, min(18, age))
    median = WHO_GUIDELINES[gender][metric][safe_age]
    if metric == "Height": low_mult, high_mult = 0.93, 1.07
    else: low_mult, high_mult = 0.85, 1.15
    return f"{median * low_mult:.1f} - {median * high_mult:.1f}"

# --- Detailed Terminology ---
def categorize_metric(metric, val, age, gender):
    if gender not in ['M', 'F'] or age is None or val is None: return "N/A" 
    safe_age = max(5, min(18, age))
    try:
        median = WHO_GUIDELINES[gender][metric][safe_age]
    except KeyError:
        return "N/A"
    
    if metric == "Height":
        low_mult, high_mult = 0.93, 1.07
        word_low, word_high = "Short", "Tall"
    else: 
        low_mult, high_mult = 0.85, 1.15
        word_low, word_high = "Underweight", "Overweight"

    lower_bound = median * low_mult
    upper_bound = median * high_mult

    if val < lower_bound:
        diff = lower_bound - val
        if diff <= 1.0: return f"Slightly {word_low}"
        elif diff <= 5.0: return word_low
        else: return f"Very {word_low}"

    elif val > upper_bound:
        diff = val - upper_bound
        if diff <= 1.0: return f"Slightly {word_high}"
        elif diff <= 5.0: return word_high
        else: return f"Very {word_high}"

    return "Normal"

def safe_float(val):
    try: return float(val)
    except ValueError: return None

class GrowthReportApp:
    def __init__(self, root):
        self.root = root
        self.session_entry_count = 0
        self.quote_index = 0
        
        if len(sys.argv) > 1:
            passed_db_path = sys.argv[1]
        else:
            self.root.withdraw() 
            messagebox.showerror("Access Denied", "Direct access is disabled.\n\nPlease launch this tool through the LVLX School Portal.")
            sys.exit() 

        raw_name = os.path.basename(passed_db_path).replace('.db', '')
        self.display_school_name = raw_name.replace('_', ' ').replace('-', ' ')

        base_reports_dir = os.path.join(os.getcwd(), "Reports")
        school_dir = os.path.join(base_reports_dir, raw_name)
        db_folder = os.path.join(school_dir, "database")
        os.makedirs(db_folder, exist_ok=True)
        
        self.db_name = os.path.join(db_folder, f"{raw_name}.db")
        
        if os.path.abspath(passed_db_path) != os.path.abspath(self.db_name):
            if os.path.exists(passed_db_path):
                shutil.move(passed_db_path, self.db_name)

        self.root.title(f"LVLX Growth Report Generator - {self.display_school_name}")
        self.root.geometry("880x750") 

        # --- CROSS-PLATFORM THEME ENGINE ---
        self.is_mac = sys.platform == 'darwin'
        self.bg_color = "#f4f6f7"
        self.root.configure(bg=self.bg_color) 

        style = ttk.Style()
        style.theme_use('clam')
        
        # Force strict color rules to bypass Mac Dark Mode rendering bugs
        style.configure('TFrame', background=self.bg_color)
        style.configure('TLabel', background=self.bg_color, foreground="black")
        style.configure('TEntry', fieldbackground="white", foreground="black", insertcolor="black")
        style.configure('TRadiobutton', background=self.bg_color, foreground="black")

        self.init_db()

        self.report_date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.create_menu()

        self.main_frame = tk.Frame(self.root, bg="#f4f6f7")
        self.main_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(self.main_frame, highlightthickness=0, bg="#f4f6f7")
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#f4f6f7", padx=10, pady=10)

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        def center_content():
            canvas_width = self.canvas.winfo_width()
            frame_width = self.scrollable_frame.winfo_reqwidth()
            x_pos = max(0, (canvas_width - frame_width) // 2)
            self.canvas.coords(self.canvas_window, x_pos, 0)
            
        self.center_content = center_content

        def on_frame_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            self.center_content()

        def on_canvas_configure(event):
            self.center_content()

        self.scrollable_frame.bind("<Configure>", on_frame_configure)
        self.canvas.bind('<Configure>', on_canvas_configure)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        def _on_mousewheel_linux_up(event):
            self.canvas.yview_scroll(-1, "units")
        def _on_mousewheel_linux_down(event):
            self.canvas.yview_scroll(1, "units")

        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.canvas.bind_all("<Button-4>", _on_mousewheel_linux_up)
        self.canvas.bind_all("<Button-5>", _on_mousewheel_linux_down)

        self.vcmd_pct = (self.root.register(self.validate_pct), '%P')

        self.erp_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.age_var = tk.StringVar()
        self.gender_var = tk.StringVar(value="")
        self.class_var = tk.StringVar()
        self.height_var = tk.StringVar()
        self.weight_var = tk.StringVar()
        self.bmi_var = tk.StringVar()
        
        self.h_obs_percentile_var = tk.StringVar(value="50")
        self.h_obs_status_var = tk.StringVar(value="")
        self.w_obs_percentile_var = tk.StringVar(value="50")
        self.w_obs_status_var = tk.StringVar(value="")
        self.overall_status_var = tk.StringVar(value="")

        self.tbl_h_val, self.tbl_h_rng, self.tbl_h_stat = tk.StringVar(value="-"), tk.StringVar(value="-"), tk.StringVar(value="-")
        self.tbl_w_val, self.tbl_w_rng, self.tbl_w_stat = tk.StringVar(value="-"), tk.StringVar(value="-"), tk.StringVar(value="-")
        self.tbl_b_val, self.tbl_b_rng, self.tbl_b_stat = tk.StringVar(value="-"), tk.StringVar(value="-"), tk.StringVar(value="-")
        
        self.row_widgets = []

        self.age_var.trace_add("write", self.update_live_metrics)
        self.gender_var.trace_add("write", self.update_live_metrics)
        self.height_var.trace_add("write", self.update_live_metrics)
        self.weight_var.trace_add("write", self.update_live_metrics)
        
        self.h_obs_percentile_var.trace_add("write", self.update_observation_text)
        self.h_obs_status_var.trace_add("write", self.update_observation_text)
        self.w_obs_percentile_var.trace_add("write", self.update_observation_text)
        self.w_obs_status_var.trace_add("write", self.update_observation_text)

        self.create_widgets()
        self.update_observation_text() 

    def get_btn_style(self, hex_color):
        """Dynamically styles buttons based on Operating System"""
        if self.is_mac:
            return {"highlightbackground": hex_color, "fg": "black"}
        else:
            return {"bg": hex_color, "fg": "white"}

    def init_db(self):
        db_filename = os.path.basename(self.db_name)
        school_name = os.path.splitext(db_filename)[0]
        base_reports_dir = os.path.join(os.getcwd(), "Reports")
        school_dir = os.path.join(base_reports_dir, school_name)

        template_dir = os.path.join(os.getcwd(), "template")
        os.makedirs(template_dir, exist_ok=True)

        required_folders = [
            "inbody_master", "inbody_report", "interpretation_docs", 
            "Master Report", "inbody_data", "student_data", "parents_data", "logs", "database", "backups",
            os.path.join("exports", "student_data"),
            os.path.join("exports", "charts_data")
        ]
        for folder in required_folders: os.makedirs(os.path.join(school_dir, folder), exist_ok=True)

        student_csv_path = os.path.join(school_dir, "student_data", "student_data.csv")
        if not os.path.exists(student_csv_path):
            with open(student_csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['ID', 'Name', 'Class'])

        parents_csv_path = os.path.join(school_dir, "parents_data", "parents_data.csv")
        if not os.path.exists(parents_csv_path):
            with open(parents_csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['ID', 'Mail', 'Number'])

        with sqlite3.connect(self.db_name) as conn:
            conn.cursor().execute('''CREATE TABLE IF NOT EXISTS growth_reports (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    name TEXT, age TEXT, gender TEXT, erp TEXT, class_grade TEXT, report_date TEXT, height REAL, weight REAL, bmi REAL,
                    observations TEXT, height_range TEXT, height_status TEXT, weight_range TEXT, weight_status TEXT, bmi_range TEXT, bmi_status TEXT)''')

    def create_menu(self):
        menubar = tk.Menu(self.root)
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="💾 Create Daily Backup", command=self.create_backup)
        tools_menu.add_separator()
        tools_menu.add_command(label="✂️ Split Master InBody PDFs", command=self.split_master_pdfs)
        tools_menu.add_separator()
        tools_menu.add_command(label="📅 Change Report Date", command=self.open_date_picker)
        tools_menu.add_separator()
        tools_menu.add_command(label="🔄 Standardize Folder Structure (Date/File)", command=self.migrate_old_reports)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        self.root.config(menu=menubar)

    def create_backup(self):
        if not messagebox.askyesno("Confirm Backup", "Do you want to create a full daily backup of this school's data?\n\nThis will securely compress all reports, documents, and databases into a single ZIP file."):
            return

        self.root.config(cursor="wait")
        self.root.update()

        try:
            db_filename = os.path.basename(self.db_name)
            school_name = os.path.splitext(db_filename)[0]
            school_dir = os.path.join(os.getcwd(), "Reports", school_name)
            backups_dir = os.path.join(school_dir, "backups")
            
            os.makedirs(backups_dir, exist_ok=True)
            
            current_date = datetime.now().strftime("%d-%m-%Y")
            zip_filename = f"{self.display_school_name} ({current_date}).zip"
            zip_filepath = os.path.join(backups_dir, zip_filename)
            
            with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root_dir, dirs, files in os.walk(school_dir):
                    if os.path.abspath(root_dir).startswith(os.path.abspath(backups_dir)):
                        continue
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path, start=school_dir)
                        zipf.write(file_path, arcname)
                        
            messagebox.showinfo("Backup Successful", f"Daily Backup Complete! 📦\n\nFile Name: {zip_filename}\nSaved in the 'backups' folder.")
        except Exception as e:
            messagebox.showerror("Backup Error", f"An error occurred while creating the backup:\n{str(e)}")
        finally:
            self.root.config(cursor="")

    def trigger_easter_egg(self):
        quotes = [
            "💧 Stay hydrated! Water is the fuel of life.",
            "🧘 Take a deep breath and stretch. You're doing great!",
            "🍏 Good health is true wealth. Keep up the awesome work!",
            "👀 Screen fatigue? Look 20 feet away for 20 seconds to rest your eyes.",
            "🪑 Posture check! Sit up straight and relax your shoulders.",
            "🌱 Cultivate your health like a garden—patience, sunshine, and some homegrown vegetables yield the best results.",
            "☕ Fuel your focus! A quick protein boost or a refreshing cold coffee can work wonders for your energy.",
            "🍲 Balance is everything. Savor that comforting bowl of noodles or soup, just remember to mix in your greens!",
            "🏃 Consistency beats intensity. Keep showing up for yourself every single day."
        ]
        
        selected_quote = quotes[self.quote_index % len(quotes)]
        self.quote_index += 1
        
        popup = tk.Toplevel(self.root)
        popup.title("✨ Milestone Reached!")
        popup.geometry("520x260")
        popup.configure(bg="#2c3e50", padx=25, pady=25)
        popup.transient(self.root)
        popup.grab_set()

        popup.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (popup.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (popup.winfo_height() // 2)
        popup.geometry(f"+{x}+{y}")

        tk.Label(popup, text=f"🎉 {self.session_entry_count} Entries Completed! 🎉", font=("Helvetica", 18, "bold"), fg="#f1c40f", bg="#2c3e50").pack(pady=(0, 10))
        tk.Label(popup, text=selected_quote, font=("Helvetica", 12, "italic"), fg="white", bg="#2c3e50", wraplength=450, justify="center").pack(pady=(0, 20))
        tk.Label(popup, text="⚠️ Friendly Reminder: Don't forget to run a Daily Backup\nonce you are done with all your entries today!", font=("Helvetica", 11, "bold"), fg="#e67e22", bg="#2c3e50").pack(pady=(0, 20))
        
        tk.Button(popup, text="Got it, thanks!", font=("Helvetica", 10, "bold"), command=popup.destroy, padx=20, pady=5, relief="groove", **self.get_btn_style("#2ecc71")).pack()

    def split_master_pdfs(self):
        try: import PyPDF2
        except ImportError: return messagebox.showerror("Missing Library", "The 'PyPDF2' library is required to run your specific splitting logic.\n\nPlease open your terminal and run:\n pip install PyPDF2")

        db_filename = os.path.basename(self.db_name)
        school_name = os.path.splitext(db_filename)[0]
        base_reports_dir = os.path.join(os.getcwd(), "Reports")
        school_dir = os.path.join(base_reports_dir, school_name)
        master_dir = os.path.join(school_dir, "inbody_master")
        output_dir = os.path.join(school_dir, "inbody_report")

        if not os.path.exists(master_dir):
            os.makedirs(master_dir, exist_ok=True)
            return messagebox.showinfo("Folder Created", f"Created a new folder at:\n{master_dir}\n\nPlease drop your bulk/multi-page InBody PDFs into this folder and run the tool again.")

        pdf_files = [os.path.join(master_dir, f) for f in os.listdir(master_dir) if f.lower().endswith('.pdf')]
        if not pdf_files: return messagebox.showwarning("No Files Found", f"No PDF files were found inside:\n{master_dir}")

        if len(pdf_files) > 15:
            display_list = "\n".join([os.path.basename(f) for f in pdf_files[:15]]) + f"\n... and {len(pdf_files)-15} more files."
        else:
            display_list = "\n".join([os.path.basename(f) for f in pdf_files])

        if not messagebox.askyesno("Confirm Split", f"Found {len(pdf_files)} Master PDF(s):\n\n{display_list}\n\nDo you want to split ALL of these into individual ERP files now?"): return

        os.makedirs(output_dir, exist_ok=True)
        
        # --- PROGRESS BAR UI SETUP ---
        progress_win = tk.Toplevel(self.root)
        progress_win.title("Splitting PDFs...")
        progress_win.geometry("420x160")
        progress_win.transient(self.root)
        progress_win.grab_set()

        progress_win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (420 // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (160 // 2)
        progress_win.geometry(f"+{x}+{y}")

        ttk.Label(progress_win, text="Processing and Splitting Master Reports...", font=("Helvetica", 11, "bold")).pack(pady=(20, 10))
        
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(progress_win, variable=progress_var, maximum=100, length=320)
        progress_bar.pack(pady=5)
        
        status_lbl = ttk.Label(progress_win, text="Calculating total pages...", font=("Helvetica", 9))
        status_lbl.pack(pady=5)

        self.root.config(cursor="wait")
        self.root.update()

        generated_count = 0
        try:
            total_pages = 0
            pdf_readers = []
            for input_file in pdf_files:
                reader = PyPDF2.PdfReader(input_file)
                total_pages += len(reader.pages)
                pdf_readers.append((input_file, reader))

            if total_pages == 0:
                progress_win.destroy()
                self.root.config(cursor="")
                return messagebox.showwarning("No Pages", "The selected PDF files appear to be empty.")

            progress_bar["maximum"] = total_pages
            current_page_idx = 0

            for input_file, pdf_reader in pdf_readers:
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    page_text = page.extract_text()
                    match = re.search(r'Test Date / Time(?:\[.*?\])?\n([A-Za-z0-9_-]+)', page_text)
                    if match: page_id = match.group(1).strip()
                    else:
                        lines = page_text.split('\n')
                        if len(lines) > 1 and 'ID' in lines[0]: page_id = lines[1].strip().split()[0]
                        else: page_id = f"unknown_id_page_{generated_count+1}"
                            
                    output_filepath = os.path.join(output_dir, f"{page_id}.pdf")
                    pdf_writer = PyPDF2.PdfWriter()
                    pdf_writer.add_page(page)
                    with open(output_filepath, 'wb') as output_file: pdf_writer.write(output_file)
                    
                    generated_count += 1
                    current_page_idx += 1
                    
                    progress_var.set(current_page_idx)
                    status_lbl.config(text=f"Exporting {page_id}.pdf ({current_page_idx}/{total_pages})")
                    progress_win.update()

            progress_win.destroy()
            messagebox.showinfo("Success", f"Extraction Complete! 🎉\n\nSuccessfully split and saved {generated_count} individual reports directly into the 'inbody_report' folder.")
            
        except Exception as e: 
            if progress_win.winfo_exists():
                progress_win.destroy()
            messagebox.showerror("Splitting Error", f"An error occurred while splitting the PDFs:\n{str(e)}")
        finally: 
            self.root.config(cursor="")

    def open_date_picker(self):
        picker = tk.Toplevel(self.root)
        picker.title("Select Date")
        picker.geometry("320x160")
        picker.transient(self.root) 
        picker.grab_set()           

        ttk.Label(picker, text="Set Active Report Date:", font=("Helvetica", 12, "bold")).pack(pady=15)
        frame = ttk.Frame(picker)
        frame.pack(pady=5)

        curr_date = self.report_date_var.get().split('/')
        day_var = tk.StringVar(value=curr_date[0])
        month_var = tk.StringVar(value=curr_date[1])
        year_var = tk.StringVar(value=curr_date[2])

        days = [f"{i:02d}" for i in range(1, 32)]
        months = [f"{i:02d}" for i in range(1, 13)]
        current_year = int(datetime.now().strftime("%Y"))
        years = [str(i) for i in range(current_year - 5, current_year + 5)]

        ttk.Combobox(frame, textvariable=day_var, values=days, width=3, state="readonly").pack(side="left", padx=2)
        ttk.Label(frame, text="/", font=("Helvetica", 12)).pack(side="left")
        ttk.Combobox(frame, textvariable=month_var, values=months, width=3, state="readonly").pack(side="left", padx=2)
        ttk.Label(frame, text="/", font=("Helvetica", 12)).pack(side="left")
        ttk.Combobox(frame, textvariable=year_var, values=years, width=5, state="readonly").pack(side="left", padx=2)

        def save_date():
            self.report_date_var.set(f"{day_var.get()}/{month_var.get()}/{year_var.get()}")
            picker.destroy()

        ttk.Button(picker, text="Set Date", command=save_date).pack(pady=15)

    def migrate_old_reports(self):
        if not messagebox.askyesno("Confirm Standardization", "This will scan your old Master Reports and Interpretation Docs, and neatly flatten them strictly by Date.\n\nDo you want to proceed?"): return

        db_filename = os.path.basename(self.db_name)
        school_name = os.path.splitext(db_filename)[0]
        base_reports_dir = os.path.join(os.getcwd(), "Reports")
        school_dir = os.path.join(base_reports_dir, school_name)
        
        moved_count = 0
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                interp_dir = os.path.join(school_dir, "interpretation_docs")
                if os.path.exists(interp_dir):
                    for root, dirs, files in os.walk(interp_dir, topdown=False):
                        for file in files:
                            if file.startswith("LVLX_Growth_Report_") and file.endswith(".docx"):
                                erp = file.replace("LVLX_Growth_Report_", "").replace(".docx", "")
                                cursor.execute("SELECT report_date FROM growth_reports WHERE erp=? ORDER BY id DESC LIMIT 1", (erp,))
                                row = cursor.fetchone()
                                if row and row[0]:
                                    date_folder = str(row[0]).replace('/', '-')
                                    target_dir = os.path.join(interp_dir, date_folder)
                                    os.makedirs(target_dir, exist_ok=True)
                                    src = os.path.join(root, file)
                                    dst = os.path.join(target_dir, file)
                                    if src != dst:
                                        shutil.move(src, dst)
                                        moved_count += 1
                        if root != interp_dir and not os.listdir(root):
                            try: os.rmdir(root)
                            except: pass

                master_dir = os.path.join(school_dir, "Master Report")
                if os.path.exists(master_dir):
                    for root, dirs, files in os.walk(master_dir, topdown=False):
                        for file in files:
                            if file.startswith("LVLX_Master_Report_") and file.endswith(".pdf"):
                                erp = file.replace("LVLX_Master_Report_", "").replace(".pdf", "")
                                cursor.execute("SELECT report_date FROM growth_reports WHERE erp=? ORDER BY id DESC LIMIT 1", (erp,))
                                row = cursor.fetchone()
                                if row and row[0]:
                                    date_folder = str(row[0]).replace('/', '-')
                                    target_dir = os.path.join(master_dir, date_folder)
                                    os.makedirs(target_dir, exist_ok=True)
                                    src = os.path.join(root, file)
                                    dst = os.path.join(target_dir, file)
                                    if src != dst:
                                        shutil.move(src, dst)
                                        moved_count += 1
                            elif file.startswith("temp_") and file.endswith(".pdf"):
                                try: os.remove(os.path.join(root, file))
                                except: pass
                        if root != master_dir and not os.listdir(root):
                            try: os.rmdir(root)
                            except: pass
                            
            if moved_count > 0: messagebox.showinfo("Migration Complete", f"Success! 🚀\n\nFlattened and standardized {moved_count} old files strictly into Date folders.")
            else: messagebox.showinfo("Migration Complete", "Your folders are already perfectly standardized!")
        except Exception as e: messagebox.showerror("Migration Error", f"An error occurred during migration:\n{str(e)}")

    def silent_pdf_convert(self, src, dest):
        if os.name == 'nt' and HAS_DOCX2PDF:
            with open(os.devnull, 'w') as devnull:
                old_stdout = sys.stdout
                old_stderr = sys.stderr
                try:
                    sys.stdout = devnull
                    sys.stderr = devnull
                    convert_to_pdf(src, dest)
                finally:
                    sys.stdout = old_stdout
                    sys.stderr = old_stderr
        elif sys.platform == 'darwin':
            # Cross-Platform: Natively run Mac MS Word via AppleScript
            import subprocess
            script = f'''
            tell application "Microsoft Word"
                open POSIX file "{os.path.abspath(src)}"
                set theActiveDoc to the active document
                save as theActiveDoc file format format PDF file name "{os.path.abspath(dest)}"
                close theActiveDoc saving no
            end tell
            '''
            subprocess.run(['osascript', '-e', script], check=True)

    def get_ordinal(self, n):
        if 11 <= (n % 100) <= 13: return f"{n}th"
        return f"{n}" + {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')

    def validate_pct(self, P):
        if P == "": return True 
        if P.isdigit(): return True 
        return False

    def create_widgets(self):
        self.scrollable_frame.columnconfigure(1, weight=1)

        header_frame = tk.Frame(self.scrollable_frame, bg="#2c3e50", padx=15, pady=10)
        header_frame.grid(row=0, column=0, columnspan=2, sticky="we", pady=(0, 15))
        
        tk.Label(header_frame, text=f"🏫 {self.display_school_name}", font=("Helvetica", 14, "bold"), fg="white", bg="#2c3e50").pack(side="left")
        date_display = tk.Label(header_frame, textvariable=self.report_date_var, font=("Helvetica", 11, "bold"), fg="#f39c12", bg="#2c3e50")
        date_display.pack(side="right")
        tk.Label(header_frame, text="Active Date: ", font=("Helvetica", 11), fg="white", bg="#2c3e50").pack(side="right")

        ttk.Label(self.scrollable_frame, text="ERP Number:").grid(row=1, column=0, sticky="w")
        erp_frame = tk.Frame(self.scrollable_frame, bg="#f4f6f7")
        erp_frame.grid(row=1, column=1, sticky="we", pady=2)
        ttk.Entry(erp_frame, textvariable=self.erp_var, width=15).pack(side="left", fill="x", expand=True)
        tk.Button(erp_frame, text="🔍 Auto-Fill from CSVs", font=("Helvetica", 9, "bold"), command=self.search_erp, **self.get_btn_style("#2980b9")).pack(side="left", padx=(10,0))

        ttk.Label(self.scrollable_frame, text="Name:").grid(row=2, column=0, sticky="w")
        ttk.Entry(self.scrollable_frame, textvariable=self.name_var, width=30).grid(row=2, column=1, sticky="we", pady=2)

        ttk.Label(self.scrollable_frame, text="Age & Gender:").grid(row=3, column=0, sticky="w")
        age_frame = tk.Frame(self.scrollable_frame, bg="#f4f6f7")
        age_frame.grid(row=3, column=1, sticky="we", pady=2)
        ttk.Entry(age_frame, textvariable=self.age_var, width=10).pack(side="left", fill="x", expand=True, padx=(0, 15))
        ttk.Radiobutton(age_frame, text="M", variable=self.gender_var, value="M").pack(side="left", padx=(0, 10))
        ttk.Radiobutton(age_frame, text="F", variable=self.gender_var, value="F").pack(side="left")

        ttk.Label(self.scrollable_frame, text="Class/Grade:").grid(row=4, column=0, sticky="w")
        ttk.Entry(self.scrollable_frame, textvariable=self.class_var, width=30).grid(row=4, column=1, sticky="we", pady=2)

        ttk.Label(self.scrollable_frame, text="Key Measurements", font=("Helvetica", 14, "bold"), foreground="#2c3e50").grid(row=5, column=0, columnspan=2, sticky="w", pady=(20, 5))

        ttk.Label(self.scrollable_frame, text="Height (cm):").grid(row=6, column=0, sticky="w")
        ttk.Entry(self.scrollable_frame, textvariable=self.height_var, width=30).grid(row=6, column=1, sticky="we", pady=2)

        ttk.Label(self.scrollable_frame, text="Weight (kg):").grid(row=7, column=0, sticky="w")
        ttk.Entry(self.scrollable_frame, textvariable=self.weight_var, width=30).grid(row=7, column=1, sticky="we", pady=2)

        ttk.Label(self.scrollable_frame, text="Calculated BMI:").grid(row=8, column=0, sticky="w")
        self.bmi_entry = ttk.Entry(self.scrollable_frame, textvariable=self.bmi_var, width=30, state="readonly")
        self.bmi_entry.grid(row=8, column=1, sticky="we", pady=2)

        table_frame = ttk.LabelFrame(self.scrollable_frame, text=" Live Health Analysis ", padding=10)
        table_frame.grid(row=9, column=0, columnspan=2, sticky="we", pady=(15, 10))
        
        inner_table = tk.Frame(table_frame, bg="#bdc3c7") 
        inner_table.pack(fill="both", expand=True)
        
        for col in range(4): inner_table.columnconfigure(col, weight=1)

        headers = ["Metric", "Value", "Healthy Range", "Status Indicator"]
        for col, head in enumerate(headers):
            tk.Label(inner_table, text=head, font=("Helvetica", 10, "bold"), fg="#34495e", bg="#ecf0f1", anchor="w", padx=10, pady=8).grid(row=0, column=col, sticky="nsew", padx=0, pady=(0, 1))

        metrics = [
            ("Height", self.tbl_h_val, self.tbl_h_rng, self.tbl_h_stat),
            ("Weight", self.tbl_w_val, self.tbl_w_rng, self.tbl_w_stat),
            ("BMI", self.tbl_b_val, self.tbl_b_rng, self.tbl_b_stat)
        ]
        
        self.row_widgets = []
        for r, (label, val_var, rng_var, stat_var) in enumerate(metrics, start=1):
            l1 = tk.Label(inner_table, text=label, font=("Helvetica", 10, "bold"), anchor="w", padx=10, pady=8, bg="#f4f6f7", fg="black")
            l1.grid(row=r, column=0, sticky="nsew", padx=0, pady=(0, 1))
            l2 = tk.Label(inner_table, textvariable=val_var, font=("Helvetica", 10), anchor="w", padx=10, pady=8, bg="#f4f6f7", fg="black")
            l2.grid(row=r, column=1, sticky="nsew", padx=0, pady=(0, 1))
            l3 = tk.Label(inner_table, textvariable=rng_var, font=("Helvetica", 10), anchor="w", padx=10, pady=8, bg="#f4f6f7", fg="black")
            l3.grid(row=r, column=2, sticky="nsew", padx=0, pady=(0, 1))
            l4 = tk.Label(inner_table, textvariable=stat_var, font=("Helvetica", 10, "bold"), width=18, anchor="w", padx=10, pady=8, bg="#f4f6f7", fg="black")
            l4.grid(row=r, column=3, sticky="nsew", padx=0, pady=(0, 1))
            
            self.row_widgets.append((l1, l2, l3, l4, stat_var))

        ttk.Label(self.scrollable_frame, text="Key Observations", font=("Helvetica", 14, "bold"), foreground="#2c3e50").grid(row=10, column=0, columnspan=2, sticky="w", pady=(20, 5))
        
        gradient_colors = ["#d4edda", "#fcf3cf", "#f7dc6f", "#f0b27a", "#e59866"]
        status_options = ["Within", "Slightly Below", "Below", "Slightly Above", "Above"]

        h_obs_frame = tk.Frame(self.scrollable_frame, bg="#f4f6f7")
        h_obs_frame.grid(row=11, column=0, columnspan=2, sticky="we", pady=(0, 5))
        
        ttk.Label(h_obs_frame, text="Height Pct:").grid(row=0, column=0, sticky="w")
        self.h_pct_entry = ttk.Entry(h_obs_frame, textvariable=self.h_obs_percentile_var, width=8, validate='key', validatecommand=self.vcmd_pct)
        self.h_pct_entry.grid(row=0, column=1, sticky="w", padx=(5, 15))
        
        ttk.Label(h_obs_frame, text="Status:").grid(row=1, column=0, sticky="w", pady=(5,0))
        h_r_frame = tk.Frame(h_obs_frame, bg="#f4f6f7")
        h_r_frame.grid(row=1, column=1, sticky="w", pady=(5,0))
        
        for idx, val in enumerate(status_options):
            tk.Radiobutton(h_r_frame, text=val, variable=self.h_obs_status_var, value=val.lower(),
                           bg=gradient_colors[idx], activebackground=gradient_colors[idx],
                           selectcolor="white", fg="black", relief="groove", borderwidth=1, tristatevalue="x").pack(side="left", padx=(0, 5), ipadx=3)

        w_obs_frame = tk.Frame(self.scrollable_frame, bg="#f4f6f7")
        w_obs_frame.grid(row=12, column=0, columnspan=2, sticky="we", pady=(10, 5))
        
        ttk.Label(w_obs_frame, text="Weight Pct:").grid(row=0, column=0, sticky="w")
        self.w_pct_entry = ttk.Entry(w_obs_frame, textvariable=self.w_obs_percentile_var, width=8, validate='key', validatecommand=self.vcmd_pct)
        self.w_pct_entry.grid(row=0, column=1, sticky="w", padx=(5, 15))
        
        ttk.Label(w_obs_frame, text="Status:").grid(row=1, column=0, sticky="w", pady=(5,0))
        w_r_frame = tk.Frame(w_obs_frame, bg="#f4f6f7")
        w_r_frame.grid(row=1, column=1, sticky="w", pady=(5,0))
        
        for idx, val in enumerate(status_options):
            tk.Radiobutton(w_r_frame, text=val, variable=self.w_obs_status_var, value=val.lower(),
                           bg=gradient_colors[idx], activebackground=gradient_colors[idx],
                           selectcolor="white", fg="black", relief="groove", borderwidth=1, tristatevalue="x").pack(side="left", padx=(0, 5), ipadx=3)

        self.obs_text = tk.Text(self.scrollable_frame, height=4, width=75, wrap="word", font=("Helvetica", 10), bg="white", fg="black", insertbackground="black", relief="solid", borderwidth=1)
        self.obs_text.grid(row=13, column=0, columnspan=2, sticky="we", pady=10)

        ttk.Label(self.scrollable_frame, text="Overall Health Status", font=("Helvetica", 14, "bold"), foreground="#2c3e50").grid(row=14, column=0, columnspan=2, sticky="w", pady=(15, 5))
        
        status_frame = tk.Frame(self.scrollable_frame, bg="#f4f6f7")
        status_frame.grid(row=15, column=0, columnspan=2, sticky="we", pady=(0, 5))
        
        tk.Radiobutton(status_frame, text="Within Range (Green)", variable=self.overall_status_var, value="green", bg="#d4edda", activebackground="#d4edda", selectcolor="white", fg="black", relief="groove", borderwidth=1, font=("Helvetica", 10, "bold"), tristatevalue="x").pack(side="left", padx=(0, 15), ipadx=10, ipady=2)
        tk.Radiobutton(status_frame, text="Borderline (Yellow)", variable=self.overall_status_var, value="yellow", bg="#fff3cd", activebackground="#fff3cd", selectcolor="white", fg="black", relief="groove", borderwidth=1, font=("Helvetica", 10, "bold"), tristatevalue="x").pack(side="left", padx=(0, 15), ipadx=10, ipady=2)
        tk.Radiobutton(status_frame, text="Needs Attention (Red)", variable=self.overall_status_var, value="red", bg="#f8d7da", activebackground="#f8d7da", selectcolor="white", fg="black", relief="groove", borderwidth=1, font=("Helvetica", 10, "bold"), tristatevalue="x").pack(side="left", padx=(0, 15), ipadx=10, ipady=2)

        self.btn_generate = tk.Button(self.scrollable_frame, text="Generate & Build Master Report", font=("Helvetica", 13, "bold"), command=self.trigger_generation, **self.get_btn_style("#2ecc71"))
        self.btn_generate.grid(row=16, column=0, columnspan=2, sticky="we", pady=(25, 20), ipady=8)

    def search_erp(self):
        erp_val = self.erp_var.get().strip()
        if not erp_val: return messagebox.showwarning("Missing Information", "Please enter an ERP number to search.")

        clean_search_erp = erp_val.upper().replace('NIS', '').replace('<', '').replace('>', '').strip()
        db_filename = os.path.basename(self.db_name)
        school_name = os.path.splitext(db_filename)[0]
        school_dir = os.path.join(os.getcwd(), "Reports", school_name)
        
        students_csv = os.path.join(school_dir, "student_data", "student_data.csv")
        inbody_data_dir = os.path.join(school_dir, "inbody_data")

        if not os.path.exists(students_csv): return messagebox.showwarning("Files Missing", f"Please ensure 'student_data.csv' is placed in:\n{os.path.dirname(students_csv)}")
        if not os.path.exists(inbody_data_dir): return messagebox.showwarning("Files Missing", f"The folder does not exist:\n{inbody_data_dir}\n\nPlease place your daily InBody CSVs inside it.")
            
        csv_files = [os.path.join(inbody_data_dir, f) for f in os.listdir(inbody_data_dir) if f.lower().endswith('.csv')]
        if not csv_files: return messagebox.showwarning("Files Missing", f"No CSV files found in:\n{inbody_data_dir}")

        csv_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        self.name_var.set("")
        self.age_var.set("")
        self.gender_var.set("")
        self.class_var.set("")
        self.height_var.set("")
        self.weight_var.set("")
        self.bmi_var.set("")

        found_student = False
        found_inbody = False
        roman_to_arabic = {'I': '1', 'II': '2', 'III': '3', 'IV': '4', 'V': '5', 'VI': '6', 'VII': '7', 'VIII': '8', 'IX': '9', 'X': '10', 'XI': '11', 'XII': '12'}

        try:
            with open(students_csv, 'r', encoding='utf-8-sig') as f:
                for row in csv.DictReader(f):
                    erp_col = next((k for k in row.keys() if k and k.strip().upper() == 'ID'), None)
                    if not erp_col: continue
                    if str(row[erp_col]).upper().replace('NIS', '').replace('<', '').replace('>', '').strip() == clean_search_erp:
                        name_col = next((k for k in row.keys() if k and 'name' in k.lower()), None)
                        if name_col: self.name_var.set(str(row[name_col]).strip())
                        class_col = next((k for k in row.keys() if k and 'class' in k.lower()), None)
                        if class_col:
                            raw_class = str(row[class_col]).upper().replace('CLASS', '').strip()
                            self.class_var.set(roman_to_arabic.get(raw_class, raw_class))
                        found_student = True
                        break
        except Exception as e: messagebox.showerror("Error", f"Error reading student data:\n{e}")

        for inbody_csv in csv_files:
            try:
                with open(inbody_csv, 'r', encoding='utf-8-sig') as f:
                    for row in csv.DictReader(f):
                        erp_col = next((k for k in row.keys() if k and '1. ID' in k), None)
                        if not erp_col: continue
                        if str(row[erp_col]).upper().replace('NIS', '').replace('<', '').replace('>', '').strip() == clean_search_erp:
                            for k, v in row.items():
                                if not k or not v: continue
                                kl = k.lower()
                                if 'height' in kl and 'limit' not in kl: self.height_var.set(str(v).strip())
                                elif 'age' in kl: self.age_var.set(str(v).strip())
                                elif 'gender' in kl:
                                    g = str(v).strip().upper()
                                    self.gender_var.set('M' if g.startswith('M') or g == 'BOY' else 'F')
                                elif 'weight' in kl and 'limit' not in kl and 'normal' not in kl and 'target' not in kl and 'control' not in kl: self.weight_var.set(str(v).strip())
                            found_inbody = True
                            break 
            except Exception as e: print(f"Skipping file {inbody_csv} due to error: {e}")
            if found_inbody: break 

        if found_student and found_inbody: pass 
        elif found_student: messagebox.showinfo("Partial Success", "Student data found, but InBody data was missing for this ERP across all CSV files in the folder.")
        elif found_inbody: messagebox.showinfo("Partial Success", "InBody data found, but Student data was missing for this ERP.")
        else: messagebox.showwarning("Not Found", "Could not find this ERP in either the student database or any of the InBody CSVs.")

    def update_live_metrics(self, *args):
        h_str, w_str = self.height_var.get().strip(), self.weight_var.get().strip()

        if ',' in h_str or ',' in w_str: self.bmi_var.set("Error: Use dot (.)")
        else:
            try:
                h, w = float(h_str) / 100, float(w_str)
                if h > 0 and w > 0: self.bmi_var.set(f"{(w / (h * h)):.1f}")
                else: self.bmi_var.set("")
            except ValueError: self.bmi_var.set("")

        age_int, gender_val = safe_float(self.age_var.get().strip()), self.gender_var.get()
        h_val, w_val, b_val = safe_float(h_str), safe_float(w_str), safe_float(self.bmi_var.get())

        self.tbl_h_val.set(f"{h_val:.1f} cm" if h_val else "-")
        self.tbl_w_val.set(f"{w_val:.1f} kg" if w_val else "-")
        self.tbl_b_val.set(f"{b_val:.1f}" if b_val else "-")

        if age_int and gender_val in ['M', 'F']:
            self.tbl_h_rng.set(get_healthy_range('Height', age_int, gender_val))
            self.tbl_w_rng.set(get_healthy_range('Weight', age_int, gender_val))
            self.tbl_b_rng.set(get_healthy_range('BMI', age_int, gender_val))
            self.tbl_h_stat.set(categorize_metric('Height', h_val, age_int, gender_val) if h_val else "-")
            self.tbl_w_stat.set(categorize_metric('Weight', w_val, age_int, gender_val) if w_val else "-")
            self.tbl_b_stat.set(categorize_metric('BMI', b_val, age_int, gender_val) if b_val else "-")
        else:
            for v in [self.tbl_h_rng, self.tbl_w_rng, self.tbl_b_rng, self.tbl_h_stat, self.tbl_w_stat, self.tbl_b_stat]: v.set("-")

        for l1, l2, l3, l4, var in self.row_widgets:
            val = var.get().lower()
            if val == "normal": 
                bg_color, fg_color = "#d4edda", "#155724" 
            elif "slightly" in val: 
                bg_color, fg_color = "#fff3cd", "#856404" 
            elif val in ["-", "n/a", ""]:
                bg_color, fg_color = "#f4f6f7", "black"    
            else:
                bg_color, fg_color = "#f8d7da", "#721c24" 
                
            for lbl in (l1, l2, l3, l4):
                lbl.config(bg=bg_color, fg=fg_color)

    def update_observation_text(self, *args):
        h_pct_raw = ''.join(filter(str.isdigit, self.h_obs_percentile_var.get()))
        w_pct_raw = ''.join(filter(str.isdigit, self.w_obs_percentile_var.get()))
        
        h_pct = self.get_ordinal(int(h_pct_raw)) if h_pct_raw else "50th"
        w_pct = self.get_ordinal(int(w_pct_raw)) if w_pct_raw else "50th"
        
        h_stat = self.h_obs_status_var.get() or "[Select Height Status]"
        w_stat = self.w_obs_status_var.get() or "[Select Weight Status]"
        
        template = f"Your child’s height falls at the {h_pct} percentile (which is {h_stat} the expected range).\nYour child’s weight falls at the {w_pct} percentile (which is {w_stat} the expected range)."
        
        self.obs_text.delete("1.0", tk.END)
        self.obs_text.insert(tk.END, template)

    def reset_form(self):
        for var in [self.name_var, self.age_var, self.erp_var, self.class_var, self.height_var, self.weight_var, self.bmi_var]: var.set("")
        self.gender_var.set("")
        self.overall_status_var.set("") 
        self.h_obs_percentile_var.set("50")
        self.h_obs_status_var.set("")
        self.w_obs_percentile_var.set("50")
        self.w_obs_status_var.set("")
        self.root.focus()
        self.canvas.yview_moveto(0) 

    def prepare_status_image(self, original_path, save_path):
        img = Image.open(original_path).convert("RGBA")
        pixels = img.load()
        
        for y in range(img.height):
            for x in range(img.width):
                r, g, b, a = pixels[x, y]
                if r > 235 and g > 235 and b > 235:
                    pixels[x, y] = (255, 255, 255, 0)

        img.save(save_path, "PNG")

    def trigger_generation(self):
        self.btn_generate.config(text="Building Master Report... Please Wait", state="disabled", **self.get_btn_style("#e67e22"))
        self.root.config(cursor="wait")
        self.root.update()
        try: self.generate_document()
        finally:
            self.btn_generate.config(text="Generate & Build Master Report", state="normal", **self.get_btn_style("#2ecc71"))
            self.root.config(cursor="")

    def generate_document(self):
        erp_str, name_str = self.erp_var.get().strip(), self.name_var.get().strip()
        age_str, class_str, gender_val = self.age_var.get().strip(), self.class_var.get().strip(), self.gender_var.get()
        h_str, w_str, overall_status = self.height_var.get().strip(), self.weight_var.get().strip(), self.overall_status_var.get()

        if not self.h_obs_status_var.get(): return messagebox.showwarning("Validation Error", "Please select a Height Status for Key Observations.")
        if not self.w_obs_status_var.get(): return messagebox.showwarning("Validation Error", "Please select a Weight Status for Key Observations.")
        if not erp_str: return messagebox.showwarning("Validation Error", "ERP Number cannot be blank.")
        if not name_str: return messagebox.showwarning("Validation Error", "Name cannot be blank.")
        if not name_str.replace(" ", "").isalpha(): return messagebox.showwarning("Validation Error", "Name must contain only letters and spaces.")
        if not age_str: return messagebox.showwarning("Validation Error", "Age cannot be blank.")
        if not age_str.isdigit(): return messagebox.showwarning("Validation Error", "Age must be a valid integer.")
        if not gender_val: return messagebox.showwarning("Validation Error", "Please select a Gender.")
        if not class_str: return messagebox.showwarning("Validation Error", "Class/Grade cannot be blank.")
        if not class_str.isdigit(): return messagebox.showwarning("Validation Error", "Class/Grade must be a valid integer.")
        if not h_str or not w_str: return messagebox.showwarning("Validation Error", "Height and Weight cannot be blank.")
        if ',' in h_str or ',' in w_str: return messagebox.showerror("Input Error", "Please use a dot (.) for decimals.")
        if not overall_status: return messagebox.showwarning("Validation Error", "Please select an Overall Health Status.")

        try: h_val, w_val = float(h_str), float(w_str)
        except ValueError: return messagebox.showerror("Input Error", "Height and Weight must be valid numbers.")

        template_dir = os.path.join(os.getcwd(), "template")
        template_path = os.path.join(template_dir, "master_template.docx")
        
        if not os.path.exists(template_path): 
            return messagebox.showerror("Error", f"Could not find '{template_path}'. Please ensure 'master_template.docx' is inside the 'template' folder.")

        try:
            doc = DocxTemplate(template_path)
            
            combined_age, combined_date = f"{age_str}/{gender_val}", self.report_date_var.get()
            
            obs_str = self.obs_text.get("1.0", "end-1c").strip()
            obs_lines = [line.strip() for line in obs_str.split('\n') if line.strip()]
            height_sentence = obs_lines[0] if len(obs_lines) > 0 else ""
            weight_sentence = obs_lines[1] if len(obs_lines) > 1 else ""
            
            age_int, b_val = int(age_str), safe_float(self.bmi_var.get())

            gender_prefix = "male" if gender_val == 'M' else "female"
            status_map = {"green": "within range", "yellow": "borderline", "red": "needs attention"}
            status_bold_map = {"green": "WITHIN RANGE", "yellow": "BORDERLINE", "red": "NEEDS ATTENTION"}
            
            status_suffix = status_map.get(overall_status, "within range")
            bold_text_val = status_bold_map.get(overall_status, "WITHIN RANGE")
            
            target_img_name = f"{gender_prefix}-{status_suffix}.png"
            img_path = os.path.join(template_dir, target_img_name)
            
            if os.path.exists(img_path):
                transparent_img_path = os.path.join(template_dir, "temp_transparent_status.png")
                
                if HAS_PILLOW:
                    try:
                        self.prepare_status_image(img_path, transparent_img_path)
                        status_image = InlineImage(doc, transparent_img_path, width=Mm(140), height=Mm(60))
                    except Exception as img_err:
                        print(f"Error handling image: {img_err}")
                        try:
                            status_image = InlineImage(doc, img_path, width=Mm(140), height=Mm(60))
                        except Exception:
                            status_image = f"[CORRUPTED IMAGE FILE: {target_img_name}]"
                else:
                    status_image = InlineImage(doc, img_path, width=Mm(140), height=Mm(60))
            else:
                status_image = f"[MISSING IMAGE: {target_img_name}]"

            rt_status = RichText('Your child is currently falling under the "')
            rt_status.add(bold_text_val, bold=True)
            rt_status.add('" category.')

            context = {
                "name": name_str, "age": combined_age, "erp": erp_str,
                "class": class_str, "date": combined_date, "height": h_str,
                "weight": w_str, "bmi": self.bmi_var.get(),
                "observations": obs_str,
                
                "healthy_height": get_healthy_range('Height', age_int, gender_val), 
                "tbl_height_status": categorize_metric('Height', h_val, age_int, gender_val),
                "healthy_weight": get_healthy_range('Weight', age_int, gender_val), 
                "tbl_weight_status": categorize_metric('Weight', w_val, age_int, gender_val),
                
                "healthy_bmi": get_healthy_range('BMI', age_int, gender_val), 
                "bmi_status": categorize_metric('BMI', b_val, age_int, gender_val),
                
                "height_status": height_sentence,
                "weight_status": weight_sentence,
                
                "status_slider": status_image,
                "status_text": rt_status
            }

            doc.render(context)

            db_filename = os.path.basename(self.db_name)
            school_name = os.path.splitext(db_filename)[0]
            date_folder = combined_date.replace('/', '-')
            
            target_dir = os.path.join(os.getcwd(), "Reports", school_name, "interpretation_docs", date_folder)
            os.makedirs(target_dir, exist_ok=True)
            save_path = os.path.join(target_dir, f"LVLX_Growth_Report_{erp_str}.docx")
            doc.save(save_path)

            if HAS_PILLOW and os.path.exists(os.path.join(template_dir, "temp_transparent_status.png")):
                try: os.remove(os.path.join(template_dir, "temp_transparent_status.png"))
                except: pass

            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM growth_reports WHERE erp = ?", (erp_str,))
                if cursor.fetchone():
                    cursor.execute('''UPDATE growth_reports SET name=?, age=?, gender=?, class_grade=?, report_date=?, height=?, weight=?, bmi=?, observations=?, height_range=?, height_status=?, weight_range=?, weight_status=?, bmi_range=?, bmi_status=? WHERE erp=?''', 
                                   (name_str, age_str, gender_val, class_str, combined_date, h_val, w_val, b_val, obs_str, context["healthy_height"], context["tbl_height_status"], context["healthy_weight"], context["tbl_weight_status"], context["healthy_bmi"], context["bmi_status"], erp_str))
                else:
                    cursor.execute('''INSERT INTO growth_reports (name, age, gender, erp, class_grade, report_date, height, weight, bmi, observations, height_range, height_status, weight_range, weight_status, bmi_range, bmi_status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
                                   (name_str, age_str, gender_val, erp_str, class_str, combined_date, h_val, w_val, b_val, obs_str, context["healthy_height"], context["tbl_height_status"], context["healthy_weight"], context["tbl_weight_status"], context["healthy_bmi"], context["bmi_status"]))

            master_msg = ""
            if HAS_DOCX2PDF and HAS_PYPDF:
                try:
                    master_date_dir = os.path.join(os.getcwd(), "Reports", school_name, "Master Report", date_folder)
                    os.makedirs(master_date_dir, exist_ok=True)
                    
                    temp_lvlx_pdf = os.path.join(master_date_dir, f"temp_lvlx_{erp_str}.pdf")
                    temp_inbody_pdf = os.path.join(master_date_dir, f"temp_inbody_{erp_str}.pdf")
                    final_pdf = os.path.join(master_date_dir, f"LVLX_Master_Report_{erp_str}.pdf")
                    
                    self.silent_pdf_convert(save_path, temp_lvlx_pdf)
                    lvlx_ready = os.path.exists(temp_lvlx_pdf)
                    
                    inbody_dir = os.path.join(os.getcwd(), "Reports", school_name, "inbody_report")
                    inbody_src = None
                    if os.path.exists(inbody_dir):
                        for f in os.listdir(inbody_dir):
                            if os.path.splitext(f)[0].lower() == erp_str.lower():
                                inbody_src = os.path.join(inbody_dir, f)
                                break
                                
                    inbody_ready = False
                    if inbody_src:
                        ext = os.path.splitext(inbody_src)[1].lower()
                        if ext == '.pdf':
                            shutil.copy2(inbody_src, temp_inbody_pdf)
                            inbody_ready = True
                        elif ext in ['.jpg', '.jpeg', '.png'] and HAS_PILLOW:
                            img = Image.open(inbody_src)
                            if img.mode == 'RGBA': img = img.convert('RGB')
                            a4_canvas = Image.new('RGB', (827, 1169), 'white')
                            ratio = min(827 / img.width, 1169 / img.height)
                            img_resized = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.Resampling.LANCZOS)
                            a4_canvas.paste(img_resized, ((827 - img_resized.width) // 2, (1169 - img_resized.height) // 2))
                            a4_canvas.save(temp_inbody_pdf, "PDF", resolution=100.0)
                            inbody_ready = True

                    if lvlx_ready or inbody_ready:
                        merger = PdfWriter()
                        lvlx_reader = None
                        target_w, target_h = 595.276, 841.89  

                        if lvlx_ready:
                            lvlx_reader = PdfReader(temp_lvlx_pdf)
                            if len(lvlx_reader.pages) > 0:
                                target_w, target_h = float(lvlx_reader.pages[0].mediabox.width), float(lvlx_reader.pages[0].mediabox.height)
                            for i in range(min(2, len(lvlx_reader.pages))): merger.add_page(lvlx_reader.pages[i])
                                
                        if inbody_ready:
                            inbody_reader = PdfReader(temp_inbody_pdf)
                            for page in inbody_reader.pages:
                                try: page.scale_to(target_w, target_h)
                                except: pass
                                merger.add_page(page)
                                
                        if lvlx_ready:
                            for i in range(2, len(lvlx_reader.pages)): merger.add_page(lvlx_reader.pages[i])
                            
                        merger.write(final_pdf)
                        merger.close()
                        master_msg = f"\n\n✓ Master Report seamlessly merged and saved at:\n{master_date_dir}"
                    
                    if os.path.exists(temp_lvlx_pdf): os.remove(temp_lvlx_pdf)
                    if os.path.exists(temp_inbody_pdf): os.remove(temp_inbody_pdf)

                except Exception as master_e: master_msg = f"\n\n⚠ Word Doc saved, but Master Report failed to build: {master_e}"
            else: master_msg = "\n\n⚠ Install 'docx2pdf', 'pypdf', and 'pillow' to enable instant Master Report building."

            messagebox.showinfo("Success", f"Word Document saved successfully!{master_msg}")
            self.reset_form()
            
            self.session_entry_count += 1
            if self.session_entry_count % 10 == 0:
                self.trigger_easter_egg()

        except Exception as e:
            import traceback
            err_msg = traceback.format_exc()
            print(err_msg)
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}\n\n(See command prompt for full details)")

if __name__ == "__main__":
    root = tk.Tk()
    app = GrowthReportApp(root)
    root.mainloop()