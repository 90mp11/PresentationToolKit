import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
import pandas as pd
import os
import main
import utilities.constants as const

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.tooltip_window = None

    def enter(self, event=None):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify='left',
                         background="#2c3e50", foreground="#ecf0f1", relief='solid', borderwidth=1,
                         font=("Roboto", 10))
        label.pack(ipadx=1)

    def leave(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("CSV Processing Tool")
        self.master.geometry('600x500')
        self.create_widgets()

    def create_widgets(self):
        # Set the style
        style = ttk.Style(self.master)
        style.theme_use('clam')

        # Modern color palette
        self.master.configure(bg='#2c3e50')
        style.configure('TFrame', background='#2c3e50')
        style.configure('TLabel', background='#2c3e50', foreground='#ecf0f1', font=('Roboto', 12))
        style.configure('TButton', font=('Roboto', 12), padding=10, background='#3498db', foreground='white', borderwidth=0)
        style.configure('TCheckbutton', background='#2c3e50', foreground='#ecf0f1', font=('Roboto', 12))
        style.configure('TCombobox', font=('Roboto', 12))
        style.map('TButton', background=[('active', '#2980b9'), ('pressed', '#1abc9c')])

        # Custom font
        self.heading_font = Font(family="Roboto", size=14)
        self.default_font = Font(family="Roboto", size=12)

        # Main frame
        self.main_frame = ttk.Frame(self.master, padding="20 20 20 20")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

        # Upload button
        self.upload_btn = ttk.Button(self.main_frame, text="Upload CSV", command=self.upload_file)
        self.upload_btn.grid(row=0, column=0, pady=20, padx=20, ipadx=20, ipady=10)

        # Options frame (initially hidden)
        self.options_container = tk.Frame(self.main_frame, bg="#2c3e50", padx=10, pady=10)
        self.options_container.grid(row=1, column=0, pady=20, padx=20, sticky=(tk.W, tk.E))
        self.options_container.columnconfigure(0, weight=1)
        self.options_container.grid_remove()

        # Options label
        self.options_label = tk.Label(self.options_container, text="Options", font=self.heading_font, bg="#2c3e50", fg="#ecf0f1")
        self.options_label.grid(row=0, column=0, pady=10)

        self.option_vars = {}
        self.options = []
        self.release_group_var = tk.StringVar()
        self.release_group_combobox = None

        # Process and Quit buttons frame
        self.buttons_frame = ttk.Frame(self.main_frame, padding="10 10 10 10")
        self.buttons_frame.grid(row=2, column=0, pady=20, padx=20, sticky=(tk.W, tk.E))
        self.buttons_frame.columnconfigure(0, weight=1)

        self.process_btn = ttk.Button(self.buttons_frame, text="Process", command=self.process_file)
        self.process_btn.grid(row=0, column=0, padx=10, ipadx=10, ipady=5, sticky=(tk.E, tk.W))

        self.quit_btn = ttk.Button(self.buttons_frame, text="QUIT", command=self.master.destroy)
        self.quit_btn.grid(row=0, column=1, padx=10, ipadx=10, ipady=5, sticky=(tk.E, tk.W))

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
                self.options_container.grid()  # Show options frame after CSV upload
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showerror("Error", "No file selected")

    def display_project_options(self):
        self.clear_options()
        project_options = [
            ("Engineering", "engineering", "Generate engineering report"),
            ("Impact", "impact", "Generate impact report"),
            ("All Impacted", "allimpacted", "Generate all impacted report"),
            ("Who", "who", "Generate 'who' report"),
            ("On Hold", "onhold", "Generate on-hold projects report"),
            ("Objective", "objective", "Generate objective report"),
            ("Projects", "projects", "Generate projects report"),
            ("Output All", "output_all", "Generate all output report"),
            ("Release", "release", "Generate release report with additional input")
        ]
        for i, (text, mode, tooltip) in enumerate(project_options):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.options_container, text=text, variable=var)
            cb.grid(row=i+1, column=0, padx=10, pady=5, sticky=tk.W)
            self.option_vars[mode] = var
            self.options.append(cb)
            ToolTip(cb, tooltip)

        # Add dropdown for Release Group
        self.release_group_label = tk.Label(self.options_container, text="Release Group", bg='#2c3e50', fg='#ecf0f1')
        self.release_group_label.grid(row=len(project_options)+1, column=0, padx=10, pady=5, sticky=tk.W)
        self.release_group_combobox = ttk.Combobox(self.options_container, textvariable=self.release_group_var, state="readonly")
        self.release_group_combobox.grid(row=len(project_options)+2, column=0, padx=10, pady=5, sticky=tk.W)
        self.update_release_group_combobox()

    def display_document_options(self):
        self.clear_options()
        document_options = [
            ("Docs", "docs", "Generate documents report"),
            ("Document Changes", "document_changes", "Generate document changes report"),
            ("Release", "release", "Generate release report with additional input")
        ]
        for i, (text, mode, tooltip) in enumerate(document_options):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.options_container, text=text, variable=var)
            cb.grid(row=i+1, column=0, padx=10, pady=5, sticky=tk.W)
            self.option_vars[mode] = var
            self.options.append(cb)
            ToolTip(cb, tooltip)

        # Add dropdown for Release Group
        self.release_group_label = tk.Label(self.options_container, text="Release Group", bg='#2c3e50', fg='#ecf0f1')
        self.release_group_label.grid(row=len(document_options)+1, column=0, padx=10, pady=5, sticky=tk.W)
        self.release_group_combobox = ttk.Combobox(self.options_container, textvariable=self.release_group_var, state="readonly")
        self.release_group_combobox.grid(row=len(document_options)+2, column=0, padx=10, pady=5, sticky=tk.W)
        self.update_release_group_combobox()

    def update_release_group_combobox(self):
        # Assuming the CSV contains a column 'Release Group' with the relevant values
        try:
            df = pd.read_csv(self.file_path)
            release_groups = df['Release Group'].dropna().unique().tolist()
            self.release_group_combobox['values'] = release_groups
        except Exception as e:
            messagebox.showerror("Error", f"Error loading release groups: {e}")

    def clear_options(self):
        for option in self.options:
            option.grid_forget()
        self.options.clear()
        if self.release_group_combobox:
            self.release_group_combobox.grid_forget()
            self.release_group_label.grid_forget()
        self.option_vars.clear()

    def process_file(self):
        if not hasattr(self, 'file_path') or not self.file_path:
            messagebox.showerror("Error", "Please upload a CSV file first")
            return

        try:
            df = pd.read_csv(self.file_path)
            const.FILE_LOCATIONS['project_csv'] = self.file_path  # Update the path in constants
            const.FILE_LOCATIONS['document_csv'] = self.file_path  # Update the path in constants

            if self.option_vars.get('engineering', tk.BooleanVar(value=False)).get():
                main.engineering_presentation()
            if self.option_vars.get('impact', tk.BooleanVar(value=False)).get():
                main.impact_presentation()
            if self.option_vars.get('allimpacted', tk.BooleanVar(value=False)).get():
                main.allimpacted_presentation()
            if self.option_vars.get('who', tk.BooleanVar(value=False)).get():
                main.who_presentation()
            if self.option_vars.get('onhold', tk.BooleanVar(value=False)).get():
                main.onhold_presentation()
            if self.option_vars.get('objective', tk.BooleanVar(value=False)).get():
                main.objective_presentation()
            if self.option_vars.get('projects', tk.BooleanVar(value=False)).get():
                main.projects_presentation()
            if self.option_vars.get('output_all', tk.BooleanVar(value=False)).get():
                main.output_all_presentation()
            if self.option_vars.get('release', tk.BooleanVar(value=False)).get():
                release_group = self.release_group_var.get()
                if release_group:
                    main.release_presentation(release_group)  # Pass the release group as an argument
                else:
                    messagebox.showerror("Error", "Please select a release group for the release report")
            if self.option_vars.get('docs', tk.BooleanVar(value=False)).get():
                main.docs_presentation()
            if self.option_vars.get('document_changes', tk.BooleanVar(value=False)).get():
                main.document_changes_presentation()

        except Exception as e:
            messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("CSV Processing Tool")
root.geometry('600x500')  # Set window size
root.configure(bg='#2c3e50')  # Set background color
style = ttk.Style(root)
style.theme_use('clam')
app = Application(master=root)
app.mainloop()
