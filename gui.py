import tkinter as tk
import tkcalendar as tkcal
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
from PIL import Image, ImageTk, UnidentifiedImageError
import os
import sys
import pandas as pd
import utilities.constants as const
import utilities.builder as bu

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
        # Prevent the label from getting focus
        label.bind("<FocusIn>", lambda e: label.focus_set())

    def leave(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self, bg='#2c3e50')
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Passive Engineering Report Generator - v2")
        self.master.geometry('800x750')
        self.output_folder = os.path.join(os.path.expanduser("~"), "Downloads")
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
        style.configure('TCheckbutton', background='#2c3e50', foreground='#ecf0f1', font=('Roboto', 12),
                        relief='flat', borderwidth=0)
        style.configure('TCombobox', font=('Roboto', 12))
        style.map('TButton', background=[('active', '#2980b9'), ('pressed', '#1abc9c')])
        style.map('TCheckbutton', background=[('active', '#2c3e50'), ('selected', '#2c3e50')],
                foreground=[('active', '#ecf0f1'), ('selected', '#ecf0f1')],
                indicatorcolor=[('selected', '#ecf0f1'), ('!selected', '#ecf0f1')])

        # Custom font
        self.title_font = Font(family="Impact", size=22)
        self.heading_font = Font(family="Roboto", size=14)
        self.default_font = Font(family="Roboto", size=12)
        self.bold_font = Font(family="Roboto", size=12, weight="bold")
        self.bold_hyperlink_font = Font(family="Roboto", size=12, weight="bold", underline=True)

        # Main frame with scrollbar
        self.canvas = tk.Canvas(self.master, bg='#2c3e50')
        self.scrollbar = ttk.Scrollbar(self.master, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, padding="20 20 20 20")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Title Label
        self.title_label = ttk.Label(self.scrollable_frame, text="Passive Engineering Report Generator - v2", font=self.title_font)
        self.title_label.grid(row=0, column=0, columnspan=3, pady=(10, 0), padx=20, sticky=tk.W)

        # Add the logo image
        self.logo_image = Image.open(resource_path("LOGO.png"))
        self.logo_image = self.logo_image.resize((154, 50), Image.LANCZOS)
        self.logo_photo = ImageTk.PhotoImage(self.logo_image)
        self.logo_label = ttk.Label(self.scrollable_frame, image=self.logo_photo)
        self.logo_label.grid(row=0, column=3, pady=(10, 0), padx=20, sticky=tk.E)

        # Upload button
        self.upload_btn = ttk.Button(self.scrollable_frame, text="Upload CSV", command=self.upload_file)
        self.upload_btn.grid(row=1, column=0, pady=10, padx=20, ipadx=10, ipady=5, sticky=tk.W)

        self.file_label = ttk.Label(self.scrollable_frame, text="No File Selected", font=self.bold_font)
        self.file_label.grid(row=1, column=1, pady=10, padx=10, sticky=tk.W)

        self.select_folder_btn = ttk.Button(self.scrollable_frame, text="Select Output Folder", command=self.select_output_folder)
        self.select_folder_btn.grid(row=2, column=0, pady=10, padx=20, ipadx=10, ipady=5, sticky=tk.W)

        self.folder_label = ttk.Label(self.scrollable_frame, text=f"{self.output_folder}", font=self.bold_hyperlink_font, cursor="hand2")
        self.folder_label.grid(row=2, column=1, pady=10, padx=10, sticky=tk.W)
        self.folder_label.bind("<Button-1>", self.open_output_folder)

        self.options_container = ttk.Frame(self.scrollable_frame, padding="10 10 10 10")
        self.options_container.grid(row=3, column=0, columnspan=4, pady=20, padx=20, sticky=tk.NW)
        self.options_container.columnconfigure(0, weight=1)
        self.options_container.grid_remove()

        # Options label
        self.options_label = ttk.Label(self.options_container, text="Options", font=self.heading_font)
        self.options_label.grid(row=0, column=0, pady=10)

        self.option_vars = {}
        self.options = []
        self.release_group_var = tk.StringVar()
        self.release_group_combobox = None

        # Impacted Areas container (initially hidden)
        self.impacted_areas_frame = ttk.Frame(self.scrollable_frame, padding="10 10 10 10")
        self.impacted_areas_frame.grid(row=3, column=1, columnspan=4, pady=20, padx=20, sticky=tk.NW)
        self.impacted_areas_frame.columnconfigure(0, weight=1)
        self.impacted_areas_frame.grid_remove()

        self.impacted_areas_label = ttk.Label(self.impacted_areas_frame, text="Impacted Areas", font=self.heading_font)
        self.impacted_areas_label.grid(row=0, column=0, pady=10, padx=10, sticky=tk.W)

        self.impacted_areas_container = ScrollableFrame(self.impacted_areas_frame)
        self.impacted_areas_container.grid(row=1, column=0, pady=10, padx=10, sticky=(tk.W, tk.E))

        self.impacted_areas_vars = {}
        self.impacted_areas_checkbuttons = []

        # Process and Quit buttons frame
        self.buttons_frame = ttk.Frame(self.scrollable_frame, padding="10 10 10 10")
        self.buttons_frame.grid(row=5, column=0, columnspan=4, pady=20, padx=0, sticky=tk.NW)

        self.process_btn = ttk.Button(self.buttons_frame, text="Process", command=self.process_file)
        self.process_btn.grid(row=0, column=0, padx=10, ipadx=10, ipady=5)

        self.quit_btn = ttk.Button(self.buttons_frame, text="QUIT", command=self.master.destroy)
        self.quit_btn.grid(row=0, column=1, padx=10, ipadx=10, ipady=5)

    def upload_file(self):
        self.file_path = filedialog.askopenfilename()
        self.clear_impacted_areas()  # Clear impacted areas when a new file is uploaded
        self.impacted_areas_frame.grid_remove()  # Hide the impacted areas frame

        # Check if contact_names_frame exists before trying to hide it
        if hasattr(self, 'contact_names_frame'):
            self.contact_names_frame.grid_remove()  # Hide the contact names frame initially
        
        # Hide the date-pickers if they exist
        if hasattr(self, 'start_date_label'):
            self.start_date_label.grid_remove()
        if hasattr(self, 'start_date_entry'):
            self.start_date_entry.grid_remove()
        if hasattr(self, 'end_date_label'):
            self.end_date_label.grid_remove()
        if hasattr(self, 'end_date_entry'):
            self.end_date_entry.grid_remove()
        
        self.options_container.grid_remove()  # Hide options frame initially

        if self.file_path:
            try:
                self.file_label.config(text=self.file_path.split("/")[-1])
                df = pd.read_csv(self.file_path)
                if 'Project Updates' in df.columns:
                    self.display_project_options()
                elif 'ContactType' in df.columns:
                    self.display_contact_options()
                    self.show_contact_names()  # Automatically show contact names if Contact Log Report is detected
                elif 'Doc Reference' in df.columns:
                    self.display_document_options()
                else:
                    messagebox.showerror("Error", "Unknown file type")
                self.options_container.grid()  # Show options frame after CSV upload
            except Exception as e:
                messagebox.showerror("Error", str(e))
                
    def select_output_folder(self):
        self.output_folder = filedialog.askdirectory()
        if self.output_folder:
            self.folder_label.config(text=self.output_folder)
        else:
            self.output_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            self.folder_label.config(text=self.output_folder)

    def open_output_folder(self, event):
        os.startfile(self.output_folder)

    def display_project_options(self):
        self.clear_options()
        project_options = [
            ("Engineering", "engineering", "Generates individual reports for each Passive Engineer - one report will be generated for each Engineer"),
            ("Impact", "impact", "Generates individual reports for each business areas impacted by upcoming PEA Projects - one report will be generated for each selected area"),
            ("All Impacted", "allimpacted", "Generates a single report that contains a view of the Projects that will impact certain teams"),
            ("Objective", "objective", "Generates a single report that contains a view of the Projects that will impact certain Objective"),
            ("Output All", "output_all", "Generates master report that contains the Person view, Objective View and Impacts View in a single report")
        ]
        for i, (text, mode, tooltip) in enumerate(project_options):
            var = tk.BooleanVar(value=False)
            if text == "Engineering":
                var = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(self.options_container, text=text, variable=var, command=self.show_impacted_areas)
            cb.grid(row=i+1, column=0, padx=10, pady=5, sticky=tk.W)
            self.option_vars[mode] = var
            self.options.append(cb)
            ToolTip(cb, tooltip)

    def display_contact_options(self):
        self.clear_options()

        contact_options = [
            ("Contact Report", "contact", "Generates Contact Log report for each selected Engineer between the dates selected")
        ]

        # Configure the main grid columns
        self.options_container.grid_columnconfigure(0, minsize=300)  # Adjust width for options column
        self.options_container.grid_columnconfigure(1, minsize=200)  # Column for "Closed By" section

        # Set row weights to prevent stretching
        for i in range(3):
            self.options_container.grid_rowconfigure(i, weight=0)

        # Row 0, Column 0: Options label
        if not hasattr(self, 'options_label'):
            self.options_label = ttk.Label(self.options_container, text="Options", font=self.heading_font)
        self.options_label.grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)

        # Row 0, Column 1: Closed By label
        if not hasattr(self, 'closed_by_label'):
            self.closed_by_label = ttk.Label(self.options_container, text="Closed By", font=self.heading_font)
        self.closed_by_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)

        # Create a frame within the "Options" column for the nested grid
        self.options_frame = ttk.Frame(self.options_container)
        self.options_frame.grid(row=1, column=0, padx=5, pady=2, sticky=tk.NW)

        # Configure the nested grid columns within the options_frame
        self.options_frame.grid_columnconfigure(0, weight=1)
        self.options_frame.grid_columnconfigure(1, weight=1)

        # Row 0 in nested grid: Contact Log Report checkbox
        for i, (text, mode, tooltip) in enumerate(contact_options):
            var = tk.BooleanVar(value=False)
            if text == "Contact Report":
                var = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(self.options_frame, text=text, variable=var)
            cb.grid(row=i+1, column=0, padx=10, pady=5, stick=tk.W)
            self.option_vars[mode] = var
            self.options.append(cb)
            ToolTip(cb, tooltip)

        # Row 1 in nested grid: Start Date label and DatePicker
        self.start_date_label = ttk.Label(self.options_frame, text="Start Date")
        self.start_date_label.grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.start_date_entry = tkcal.DateEntry(self.options_frame, width=12, background='darkblue', foreground='white', borderwidth=2, year=2024, month=1, day=1)
        self.start_date_entry.grid(row=3, column=1, padx=5, pady=2, sticky=tk.W)

        # Row 2 in nested grid: End Date label and DatePicker
        self.end_date_label = ttk.Label(self.options_frame, text="End Date")
        self.end_date_label.grid(row=4, column=0, padx=5, pady=2, sticky=tk.W)
        self.end_date_entry = tkcal.DateEntry(self.options_frame, width=12, background='darkblue', foreground='white', borderwidth=2, year=2024, month=12, day=31)
        self.end_date_entry.grid(row=4, column=1, padx=5, pady=2, sticky=tk.W)

        # Row 1-2, Column 1: "Closed By" checkboxes, aligned to the top-left
        self.contact_names_frame = ttk.Frame(self.options_container, padding="5 5 5 5")
        self.contact_names_frame.grid(row=1, column=1, rowspan=3, pady=2, padx=5, sticky=tk.NW)

        self.contact_names_container = ScrollableFrame(self.contact_names_frame)
        self.contact_names_container.grid(row=0, column=0, pady=5, padx=5, sticky=(tk.W, tk.E))

        self.contact_names_vars = {}
        self.contact_names_checkbuttons = []

    def show_contact_names(self):
        self.contact_names_frame.grid()  # Ensure the contact names frame is shown
        self.update_contact_names_checkbuttons()

    def update_contact_names_checkbuttons(self):
        self.clear_contact_names()
        try:
            df = pd.read_csv(self.file_path)
            unique_names = df['Closed by'].dropna().unique()
            engineers = ["Chris Kelly", "Andy Oxford", "Luke Phillips", "Matthew Harbord", "Neil Griffin", "Gordon Pyrah", "Tom Wright"]

            for i, name in enumerate(sorted(unique_names)):
                # Default to True if the name is in the engineers list
                var = tk.BooleanVar(value=name in engineers)
                cb = ttk.Checkbutton(self.contact_names_container.scrollable_frame, text=name, variable=var)
                cb.grid(row=i, column=0, padx=10, pady=5, sticky=tk.W)
                self.contact_names_vars[name] = var
                self.contact_names_checkbuttons.append(cb)
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Closed By names: {e}")

    def clear_contact_names(self):
        for cb in self.contact_names_checkbuttons:
            cb.grid_forget()
        self.contact_names_checkbuttons.clear()
        self.contact_names_vars.clear()

    def display_document_options(self):
        self.clear_options()
        document_options = [
            ("Internal Release Board Report", "internal_release", "Generates the Internal Release Report"),
            ("Release", "release", "Generates the Document Release Report")
        ]
        for i, (text, mode, tooltip) in enumerate(document_options):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.options_container, text=text, variable=var)
            cb.grid(row=i+1, column=0, padx=10, pady=5, sticky=tk.W)
            self.option_vars[mode] = var
            self.options.append(cb)
            ToolTip(cb, tooltip)

        # Add dropdown for Release Group
        self.release_group_label = ttk.Label(self.options_container, text="Release Group")
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

    def show_impacted_areas(self):
        if self.option_vars.get('impact', tk.BooleanVar(value=False)).get():
            self.impacted_areas_frame.grid()
            self.update_impacted_areas_checkbuttons()
        else:
            self.impacted_areas_frame.grid_remove()

    def update_impacted_areas_checkbuttons(self):
        self.clear_impacted_areas()
        try:
            df = pd.read_csv(self.file_path)
            impacted_areas_series = df['Impacted Teams'].dropna().apply(lambda x: x.split(','))
            impacted_areas = sorted(set(area.strip().strip('[]"') for sublist in impacted_areas_series for area in sublist))
            for i, area in enumerate(impacted_areas):
                var = tk.BooleanVar(value=False)
                cb = ttk.Checkbutton(self.impacted_areas_container.scrollable_frame, text=area, variable=var)
                cb.grid(row=i, column=0, padx=10, pady=5, sticky=tk.W)
                self.impacted_areas_vars[area] = var
                self.impacted_areas_checkbuttons.append(cb)
        except Exception as e:
            messagebox.showerror("Error", f"Error loading impacted areas: {e}")

    def clear_options(self):
        for option in self.options:
            option.grid_forget()
        self.options.clear()
        if self.release_group_combobox:
            self.release_group_combobox.grid_forget()
            self.release_group_label.grid_forget()
        self.option_vars.clear()

        # Hide the Closed By label if it exists
        if hasattr(self, 'closed_by_label'):
            self.closed_by_label.grid_remove()

    def clear_impacted_areas(self):
        for cb in self.impacted_areas_checkbuttons:
            cb.grid_forget()
        self.impacted_areas_checkbuttons.clear()
        self.impacted_areas_vars.clear()

    def process_file(self):
        if not hasattr(self, 'file_path') or not self.file_path:
            self.show_png_image(resource_path("magic.png"))  # Path to your .png image
            return

        try:
            df = pd.read_csv(self.file_path)
            const.FILE_LOCATIONS['project_csv'] = self.file_path  # Update the path in constants
            const.FILE_LOCATIONS['document_csv'] = self.file_path  # Update the path in constants

            files_created = 0

            if self.option_vars.get('engineering', tk.BooleanVar(value=False)).get():
                files_created += bu.engineering_presentation(self.file_path, self.output_folder)
            if self.option_vars.get('impact', tk.BooleanVar(value=False)).get():
                selected_impacted_areas = [area for area, var in self.impacted_areas_vars.items() if var.get()]
                if selected_impacted_areas:
                    files_created += bu.impact_presentation(self.file_path, selected_impacted_areas, self.output_folder)
                else:
                    bu.allimpacted_presentation(self.file_path, self.output_folder)
                    files_created += 1
            if self.option_vars.get('allimpacted', tk.BooleanVar(value=False)).get():
                bu.allimpacted_presentation(self.file_path, self.output_folder)
                files_created += 1
            if self.option_vars.get('who', tk.BooleanVar(value=False)).get():
                bu.who_presentation(self.file_path, None, self.output_folder)
                files_created += 1
            if self.option_vars.get('onhold', tk.BooleanVar(value=False)).get():
                bu.onhold_presentation(self.file_path, self.output_folder)
                files_created += 1
            if self.option_vars.get('objective', tk.BooleanVar(value=False)).get():
                bu.objective_presentation(self.file_path, self.output_folder)
                files_created += 1
            if self.option_vars.get('projects', tk.BooleanVar(value=False)).get():
                bu.projects_presentation(self.file_path, self.output_folder)
                files_created += 1
            if self.option_vars.get('output_all', tk.BooleanVar(value=False)).get():
                bu.output_all_presentation(self.file_path, self.output_folder)
                files_created += 1
            if self.option_vars.get('release', tk.BooleanVar(value=False)).get():
                release_group = self.release_group_var.get()
                if release_group:
                    bu.release_presentation(self.file_path, release_group, internal=False, output_folder=self.output_folder)
                    files_created += 1
                else:
                    messagebox.showerror("Error", "Please select a release group for the release report")
            if self.option_vars.get('internal_release', tk.BooleanVar(value=False)).get():
                release_group = self.release_group_var.get()
                if release_group:
                    bu.release_presentation(self.file_path, release_group, internal=True, output_folder=self.output_folder)
                    files_created += 1
                else:
                    messagebox.showerror("Error", "Please select a release group for the internal release report")
            if self.option_vars.get('docs', tk.BooleanVar(value=False)).get():
                bu.docs_presentation(self.file_path, self.output_folder)
                files_created += 1
            if self.option_vars.get('document_changes', tk.BooleanVar(value=False)).get():
                bu.document_changes_presentation(self.file_path, self.output_folder)
                files_created += 1
            
            if self.option_vars.get('contact', tk.BooleanVar(value=False)).get():
                selected_names = [name for name, var in self.contact_names_vars.items() if var.get()]
                start_date = self.start_date_entry.get_date().strftime('%Y-%m-%d')
                end_date = self.end_date_entry.get_date().strftime('%Y-%m-%d')
                if selected_names:
                    bu.contact_report_presentation(self.file_path, self.output_folder, selected_names, start_date=start_date, end_date=end_date)
                    files_created += 1
                else:
                    bu.contact_report_presentation(self.file_path, self.output_folder, start_date=start_date, end_date=end_date)
                    files_created += 1

            # Show toast notification
            if files_created > 0:
                self.show_toast(f"Processing complete. {files_created} files saved to {self.output_folder}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def show_png_image(self, png_path):
        try:
            png_image = Image.open(png_path)
            png_photo = ImageTk.PhotoImage(png_image)
            png_window = tk.Toplevel(self)
            png_window.title("Ah ah ah - you didn't select a CSV File")
            label = tk.Label(png_window, image=png_photo)
            label.image = png_photo  # Keep a reference to avoid garbage collection
            label.pack()
            png_window.geometry(f"{png_image.width}x{png_image.height}")
        except UnidentifiedImageError:
            messagebox.showerror("Error", "Could not open the .png image.")

    def show_toast(self, message):
        toast = tk.Toplevel(self)
        toast.wm_overrideredirect(True)
        toast.attributes("-topmost", True)
        x = self.master.winfo_rootx() + 50
        y = self.master.winfo_rooty() + 50
        toast.geometry(f"+{x}+{y}")
        label = tk.Label(toast, text=message, bg="black", fg="white", font=("Roboto", 10))
        label.pack(ipadx=10, ipady=5)
        toast.after(3000, toast.destroy)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
            
    return os.path.join(base_path, relative_path)

def start_gui():
    root = tk.Tk()
    root.title("Passive Engineering Report Generator")
    root.geometry('850x680')  # Set window size
    root.configure(bg='#2c3e50')  # Set background color
    style = ttk.Style(root)
    style.theme_use('clam')
    app = Application(master=root)
    app.mainloop()

if __name__ == "__main__":
    start_gui()
