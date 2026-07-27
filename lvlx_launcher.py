import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import subprocess

# --- MAC .APP DIRECTORY FIX ---
if getattr(sys, 'frozen', False):
    if sys.platform == 'darwin' and '.app' in sys.executable:
        app_dir = os.path.dirname(os.path.dirname(os.path.dirname(sys.executable)))
        os.chdir(os.path.dirname(app_dir))
    else:
        os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

class LVLXLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("LVLX School Portal")
        self.root.geometry("450x350")
        
        # --- CROSS-PLATFORM THEME ENGINE ---
        self.is_mac = sys.platform == 'darwin'
        self.bg_color = "#ECECEC" if self.is_mac else "#f4f6f7"
        self.root.configure(padx=20, pady=20, bg=self.bg_color)

        style = ttk.Style()
        style.theme_use('clam')
        
        # Force strict color rules to bypass Mac Dark Mode rendering bugs
        style.configure('TFrame', background=self.bg_color)
        style.configure('TLabel', background=self.bg_color, foreground="black")
        style.configure('TEntry', fieldbackground="white", foreground="black", insertcolor="black")
        style.configure('TCombobox', fieldbackground="white", foreground="black", insertcolor="black")

        self.legacy_db_folder = "Databases"
        self.reports_folder = "Reports"
        
        if not os.path.exists(self.reports_folder):
            os.makedirs(self.reports_folder)

        self.selected_db_var = tk.StringVar()
        self.new_school_var = tk.StringVar()

        self.create_widgets()
        self.refresh_school_list()

    def get_btn_style(self, hex_color):
        """Dynamically styles buttons based on Operating System"""
        if self.is_mac:
            return {"highlightbackground": hex_color, "fg": "black"}
        else:
            return {"bg": hex_color, "fg": "white"}

    def create_widgets(self):
        ttk.Label(self.root, text="LVLX Ecosystem Launcher", font=("Helvetica", 16, "bold")).pack(pady=(0, 20))

        ttk.Label(self.root, text="Select Existing School:", font=("Helvetica", 10, "bold")).pack(anchor="w")
        
        db_frame = ttk.Frame(self.root)
        db_frame.pack(fill="x", pady=(5, 15))
        
        self.school_combo = ttk.Combobox(db_frame, textvariable=self.selected_db_var, state="readonly", width=30)
        self.school_combo.pack(side="left", fill="x", expand=True)
        
        tk.Button(db_frame, text="🔄", command=self.refresh_school_list).pack(side="left", padx=(5, 0))

        ttk.Label(self.root, text="OR Create New School Database:", font=("Helvetica", 10, "bold")).pack(anchor="w")
        ttk.Entry(self.root, textvariable=self.new_school_var, width=35).pack(anchor="w", pady=(5, 20))
        ttk.Label(self.root, text="(e.g., 'Oakridge_High' - avoids spaces)", font=("Helvetica", 8, "italic")).pack(anchor="w", pady=(0, 15))

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", pady=10)

        tk.Button(btn_frame, text="📝 Open Data Generator", font=("Helvetica", 11, "bold"), 
                  command=lambda: self.launch_app("lvlx_generator.py"), **self.get_btn_style("#3498db")).pack(side="left", expand=True, fill="x", padx=(0, 5), ipady=5)
        
        tk.Button(btn_frame, text="📊 Open Command Center", font=("Helvetica", 11, "bold"), 
                  command=lambda: self.launch_app("lvlx_command_center.py"), **self.get_btn_style("#9b59b6")).pack(side="right", expand=True, fill="x", padx=(5, 0), ipady=5)

    def refresh_school_list(self):
        display_names = set()
        
        if os.path.exists(self.reports_folder):
            for school_name in os.listdir(self.reports_folder):
                db_path = os.path.join(self.reports_folder, school_name, "database", f"{school_name}.db")
                if os.path.exists(db_path):
                    display_names.add(school_name)
                    
        if os.path.exists(self.legacy_db_folder):
            for f in os.listdir(self.legacy_db_folder):
                if f.endswith('.db'):
                    display_names.add(f[:-3])
        
        sorted_names = sorted(list(display_names))
        self.school_combo['values'] = sorted_names
        
        if sorted_names:
            self.school_combo.set(sorted_names[0])
        else:
            self.school_combo.set("No schools found")

    def launch_app(self, script_name):
        new_school = self.new_school_var.get().strip()
        existing_school = self.selected_db_var.get().strip()

        if not new_school and (not existing_school or existing_school == "No schools found" or existing_school == "Select a school"):
            messagebox.showwarning("Validation Error", "Please select a school from the dropdown or type a new school name.")
            return

        db_path = ""

        if new_school:
            confirm = messagebox.askyesno("Confirm New School", f"Are you sure you want to create a new database for:\n\n'{new_school}'?")
            if not confirm:
                return 

            safe_name = new_school.replace(" ", "_").replace("/", "-")
            db_path = os.path.join(self.reports_folder, safe_name, "database", f"{safe_name}.db")
            os.makedirs(os.path.dirname(db_path), exist_ok=True)
            db_path = os.path.abspath(db_path)
        else:
            legacy_path = os.path.join(self.legacy_db_folder, f"{existing_school}.db")
            new_path = os.path.join(self.reports_folder, existing_school, "database", f"{existing_school}.db")
            
            if os.path.exists(legacy_path):
                db_path = os.path.abspath(legacy_path)
            else:
                db_path = os.path.abspath(new_path)

        exe_name = os.path.join("SystemFiles", script_name.replace('.py', '.exe'))
        mac_app_name = script_name.replace('.py', '.app')

        try:
            if os.name == 'nt' and os.path.exists(exe_name):
                subprocess.Popen([exe_name, db_path])
            elif sys.platform == 'darwin' and os.path.exists(mac_app_name):
                subprocess.Popen(["open", "-n", "-a", os.path.abspath(mac_app_name), "--args", db_path])
            elif os.path.exists(script_name):
                subprocess.Popen([sys.executable, script_name, db_path])
            else:
                messagebox.showerror("Error", f"Could not find {script_name} or {mac_app_name} in this folder.")
                return

            self.root.destroy()
            
        except Exception as e:
            messagebox.showerror("Launch Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = LVLXLauncher(root)
    root.mainloop()