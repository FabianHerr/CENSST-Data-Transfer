import os
import tkinter as tk
from tkinter import filedialog, messagebox
from processor import run_transfer_logic, run_transfer_all_periods

# Suppress Tk deprecation warning on macOS
os.environ['TK_SILENCE_DEPRECATION'] = '1'

# Try to import sv_ttk for Windows 11 modern look, but don't crash on macOS
SV_TTK_AVAILABLE = False
try:
    import sv_ttk
    SV_TTK_AVAILABLE = True
except (ImportError, Exception):
    # Catch ImportError or any other error during import
    SV_TTK_AVAILABLE = False

class CNESSTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CNESST Data Transfer Tool")
        self.root.geometry("650x600")
        
        # Apply modern theme if available (Windows 11)
        # Catch TclError in case Tk version is too old (macOS system Tk 8.5.9 vs required 8.6+)
        if SV_TTK_AVAILABLE and 'sv_ttk' in globals():
            try:
                sv_ttk.set_theme("light")
            except (tk.TclError, AttributeError, Exception):
                # Silently fail if theme can't be applied (old Tk version or other issues)
                pass
        
        # Variables to store paths
        self.source_path = tk.StringVar()
        self.target_path = tk.StringVar()
        self.selected_period = tk.StringVar(value="1")

        # Main container with padding
        main_container = tk.Frame(root, padx=20, pady=20)
        main_container.pack(fill="both", expand=True)

        # --- STEP 1: HONORAIRES FOLDER ---
        step1_frame = tk.LabelFrame(main_container, text="STEP 1: SOURCE (Honoraires Folder)", 
                                    font=("Arial", 13, "bold"), padx=15, pady=15, fg="blue",
                                    relief=tk.RAISED, borderwidth=2)
        step1_frame.pack(fill="x", pady=15)
        
        tk.Label(step1_frame, text="Choose the folder containing Honoraires Excel files:",
                font=("Arial", 10)).pack(anchor="w", pady=(0, 10))
        
        source_btn = tk.Button(step1_frame, text="SELECT HONORAIRES FOLDER", 
                              bg="#4CAF50", fg="white", font=("Arial", 12, "bold"),
                              height=2, command=self.get_source_folder, cursor="hand2")
        source_btn.pack(fill="x", pady=5)
        
        self.source_label = tk.Label(step1_frame, textvariable=self.source_path, 
                                     fg="green", wraplength=550, font=("Arial", 9),
                                     justify="left", anchor="w")
        self.source_label.pack(anchor="w", pady=(5, 0))

        # --- STEP 2: SUIVI ---
        step2_frame = tk.LabelFrame(main_container, text="STEP 2: DESTINATION (Suivi)", 
                                    font=("Arial", 13, "bold"), padx=15, pady=15, fg="blue",
                                    relief=tk.RAISED, borderwidth=2)
        step2_frame.pack(fill="x", pady=15)
        
        tk.Label(step2_frame, text="Choose your master 'SUIVI' (.xlsm) file to copy TO:",
                font=("Arial", 10)).pack(anchor="w", pady=(0, 10))
        
        target_btn = tk.Button(step2_frame, text="SELECT SUIVI MASTER FILE", 
                              bg="#2196F3", fg="white", font=("Arial", 12, "bold"),
                              height=2, command=self.get_target_file, cursor="hand2")
        target_btn.pack(fill="x", pady=5)
        
        self.target_label = tk.Label(step2_frame, textvariable=self.target_path, 
                                     fg="green", wraplength=550, font=("Arial", 9),
                                     justify="left", anchor="w")
        self.target_label.pack(anchor="w", pady=(5, 0))

        # --- STEP 3: PERIOD ---
        step3_frame = tk.LabelFrame(main_container, text="STEP 3: SELECT PERIOD", 
                                    font=("Arial", 13, "bold"), padx=15, pady=15, fg="blue",
                                    relief=tk.RAISED, borderwidth=2)
        step3_frame.pack(fill="x", pady=15)
        
        period_container = tk.Frame(step3_frame)
        period_container.pack(fill="x")
        
        tk.Label(period_container, text="Choose period:",
                font=("Arial", 10)).pack(side=tk.LEFT, padx=(0, 15))
        
        self.period_menu = tk.OptionMenu(period_container, self.selected_period, 
                                         *[str(i) for i in range(1, 27)])
        self.period_menu.config(font=("Arial", 11, "bold"), width=5)
        self.period_menu.pack(side=tk.LEFT)

        # --- EXECUTE BUTTONS ---
        button_container = tk.Frame(main_container)
        button_container.pack(pady=20, fill="x")
        
        execute_btn = tk.Button(button_container, text="TRANSFER SELECTED PERIOD", 
                               bg="#27ae60", fg="white", font=("Arial", 14, "bold"),
                               height=2, command=self.execute, cursor="hand2",
                               relief=tk.RAISED, borderwidth=3)
        execute_btn.pack(pady=(0, 10), fill="x")
        
        execute_all_btn = tk.Button(button_container, text="TRANSFER ALL PERIODS", 
                                    bg="#e67e22", fg="white", font=("Arial", 14, "bold"),
                                    height=2, command=self.execute_all_periods, cursor="hand2",
                                    relief=tk.RAISED, borderwidth=3)
        execute_all_btn.pack(fill="x")

    def get_source_folder(self):
        """Select folder containing Honoraires Excel files"""
        folder = filedialog.askdirectory(
            title="Select Folder Containing Honoraires Files"
        )
        if folder:
            # Normalize path for Windows (handles spaces and special characters)
            normalized_folder = os.path.normpath(folder)
            self.source_path.set(normalized_folder)
            # Update label to show wrapped path
            self.source_label.config(text=normalized_folder, wraplength=550)
            print(f"DEBUG: Dossier source sélectionné: {normalized_folder}")

    def get_target_file(self):
        """Select target SUIVI master file with macOS-compatible filetypes"""
        file = filedialog.askopenfilename(
            title="Select Master Suivi File",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        if file:
            # Normalize path for Windows (handles spaces and special characters)
            normalized_file = os.path.normpath(file)
            self.target_path.set(normalized_file)
            # Update label to show wrapped path
            self.target_label.config(text=normalized_file, wraplength=550)
            print(f"DEBUG: Fichier Suivi sélectionné: {normalized_file}")

    def execute(self):
        """Execute the data transfer process for selected period"""
        if not self.source_path.get() or not self.target_path.get():
            messagebox.showwarning("Missing Selection", 
                                 "Please select both the folder and the Suivi file before transferring.")
            return
        
        # Normalize paths before processing
        source_folder = os.path.normpath(self.source_path.get())
        target_file = os.path.normpath(self.target_path.get())
        
        # Verify paths exist
        if not os.path.exists(source_folder):
            messagebox.showerror("Error", 
                               f"Le dossier source n'existe pas:\n{source_folder}\n\n"
                               "Veuillez sélectionner un dossier valide.")
            return
        
        if not os.path.exists(target_file):
            messagebox.showerror("Error", 
                               f"Le fichier Suivi n'existe pas:\n{target_file}\n\n"
                               "Veuillez sélectionner un fichier valide.")
            return
        
        print(f"DEBUG: Démarrage du transfert")
        print(f"DEBUG: Dossier source: {source_folder}")
        print(f"DEBUG: Fichier cible: {target_file}")
        print(f"DEBUG: Période sélectionnée: {self.selected_period.get()}")
        
        # Disable button during processing
        self.root.config(cursor="wait")
        try:
            result = run_transfer_logic(
                source_folder, 
                target_file, 
                self.selected_period.get()
            )
            messagebox.showinfo("Transfer Result", result)
        except Exception as e:
            error_msg = f"Une erreur s'est produite:\n{str(e)}"
            print(f"DEBUG: ERREUR - {error_msg}")
            messagebox.showerror("Error", error_msg)
        finally:
            self.root.config(cursor="")
    
    def execute_all_periods(self):
        """Execute the data transfer process for all periods"""
        if not self.source_path.get() or not self.target_path.get():
            messagebox.showwarning("Missing Selection", 
                                 "Please select both the folder and the Suivi file before transferring.")
            return
        
        # Normalize paths before processing
        source_folder = os.path.normpath(self.source_path.get())
        target_file = os.path.normpath(self.target_path.get())
        
        # Verify paths exist
        if not os.path.exists(source_folder):
            messagebox.showerror("Error", 
                               f"Le dossier source n'existe pas:\n{source_folder}\n\n"
                               "Veuillez sélectionner un dossier valide.")
            return
        
        if not os.path.exists(target_file):
            messagebox.showerror("Error", 
                               f"Le fichier Suivi n'existe pas:\n{target_file}\n\n"
                               "Veuillez sélectionner un fichier valide.")
            return
        
        # Confirm with user
        confirm = messagebox.askyesno("Confirm Transfer", 
                                     "This will transfer data for ALL periods (1-26) that have data.\n\n"
                                     "This may take a while. Continue?")
        if not confirm:
            return
        
        print(f"DEBUG: Démarrage du transfert pour toutes les périodes")
        print(f"DEBUG: Dossier source: {source_folder}")
        print(f"DEBUG: Fichier cible: {target_file}")
        
        # Disable button during processing
        self.root.config(cursor="wait")
        try:
            result = run_transfer_all_periods(
                source_folder, 
                target_file
            )
            messagebox.showinfo("Transfer Result", result)
        except Exception as e:
            error_msg = f"Une erreur s'est produite:\n{str(e)}"
            print(f"DEBUG: ERREUR - {error_msg}")
            messagebox.showerror("Error", error_msg)
        finally:
            self.root.config(cursor="")

if __name__ == "__main__":
    root = tk.Tk()
    app = CNESSTApp(root)
    root.mainloop()