import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import main
import utilities.constants as const

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack(padx=10, pady=10)
        self.create_widgets()

    def create_widgets(self):
        self.upload_frame = ttk.Frame(self)
        self.upload_frame.pack(pady=10)

        self.upload_btn = ttk.Button(self.upload_frame, text="Upload CSV", command=self.upload_file)
        self.upload_btn.pack(pady=5)

        self.options_frame = ttk.LabelFrame(self, text="Options")
        self.options_frame.pack(pady=10, fill="both", expand="yes")

        self.option_var = tk.StringVar(value="none")
        self.options = []

        self.process_btn = ttk.Button(self, text="Process", command=self.process_file)
        self.process_btn.pack(pady=5)

        self.quit = ttk.Button(self, text="QUIT", command=self.master.destroy)
        self.quit.pack(pady=5)

    def upload_file(self):
        self.file_path = filedialog.askopenfilename()
        if self.file_path:
            try:
                df = pd.read_csv(self.file_path)
                if 'Project Updates' in df.columns:
                    self.display_project_options()
                elif 'Doc Reference' in df.columns:
                    self.display_document_options()
                else:
                    messagebox.showerror("Error", "Unknown file type")
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showerror("Error", "No file selected")

    def display_project_options(self):
        self.clear_options()
        project_options = [
            ("Engineering", "engineering"),
            ("Impact", "impact"),
            ("All Impacted", "allimpacted"),
            ("Who", "who"),
            ("On Hold", "onhold"),
            ("Objective", "objective"),
            ("Projects", "projects"),
            ("Output All", "output_all"),
            ("Release", "release")
        ]
        for text, mode in project_options:
            b = ttk.Radiobutton(self.options_frame, text=text, variable=self.option_var, value=mode)
            b.pack(anchor="w", padx=10, pady=2)
            self.options.append(b)

    def display_document_options(self):
        self.clear_options()
        document_options = [
            ("Docs", "docs"),
            ("Document Changes", "document_changes"),
            ("Release", "release")
        ]
        for text, mode in document_options:
            b = ttk.Radiobutton(self.options_frame, text=text, variable=self.option_var, value=mode)
            b.pack(anchor="w", padx=10, pady=2)
            self.options.append(b)

    def clear_options(self):
        for option in self.options:
            option.destroy()
        self.options.clear()

    def process_file(self):
        if not hasattr(self, 'file_path') or not self.file_path:
            messagebox.showerror("Error", "Please upload a CSV file first")
            return
        
        try:
            df = pd.read_csv(self.file_path)
            const.FILE_LOCATIONS['project_csv'] = self.file_path  # Update the path in constants
            const.FILE_LOCATIONS['document_csv'] = self.file_path  # Update the path in constants
            option = self.option_var.get()

            if option == "engineering":
                main.engineering_presentation()
            elif option == "impact":
                main.impact_presentation()
            elif option == "allimpacted":
                main.allimpacted_presentation()
            elif option == "who":
                main.who_presentation()
            elif option == "onhold":
                main.onhold_presentation()
            elif option == "objective":
                main.objective_presentation()
            elif option == "projects":
                main.projects_presentation()
            elif option == "docs":
                main.docs_presentation()
            elif option == "document_changes":
                main.document_changes_presentation()
            elif option == "output_all":
                main.output_all_presentation()
            elif option == "release":
                main.release_presentation()
            else:
                messagebox.showerror("Error", "Please select an option to process the file")
        except Exception as e:
            messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("CSV Processing Tool")
style = ttk.Style(root)
style.theme_use('clam')
app = Application(master=root)
app.mainloop()
