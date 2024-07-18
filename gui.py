import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
import pandas as pd
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
        self.master.title("CSV Processing Tool")
        self.master.geometry('800x600')
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
        self.heading_font = Font(family="Roboto", size=14)
        self.default_font = Font(family="Roboto", size=12)

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

        # Upload button
        self.upload_btn = ttk.Button(self.scrollable_frame, text="Upload CSV", command=self.upload_file)
        self.upload_btn.grid(row=0, column=0, pady=20, padx=20, ipadx=20, ipady=10, sticky=tk.W)

        self.select_folder_btn = ttk.Button(self.scrollable_frame, text="Select Output Folder", command=self.select_output_folder)
        self.select_folder_btn.grid(row=0, column=1, pady=20, padx=20, ipadx=20, ipady=10, sticky=tk.W)

        self.options_container = ttk.Frame(self.scrollable_frame, padding="10 10 10 10")
        self.options_container.grid(row=1, column=0, pady=20, padx=20, sticky=tk.W)
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
        self.impacted_areas_frame.grid(row=1, column=1, pady=20, padx=20, sticky=tk.NW)
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
        self.buttons_frame.grid(row=2, column=0, pady=20, padx=20, sticky=tk.W)
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

    def select_output_folder(self):
        self.output_folder = filedialog.askdirectory()
        if self.output_folder:
            print(f"Selected output folder: {self.output_folder}")
        else:
            self.output_folder = const.FILE_LOCATIONS['output_folder']
            print(f"No folder selected, using default: {self.output_folder}")

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
            cb = ttk.Checkbutton(self.options_container, text=text, variable=var, command=self.show_impacted_areas)
            cb.grid(row=i+1, column=0, padx=10, pady=5, sticky=tk.W)
            self.option_vars[mode] = var
            self.options.append(cb)
            ToolTip(cb, tooltip)

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

    def clear_impacted_areas(self):
        for cb in self.impacted_areas_checkbuttons:
            cb.grid_forget()
        self.impacted_areas_checkbuttons.clear()
        self.impacted_areas_vars.clear()

    def process_file(self):
        if not hasattr(self, 'file_path') or not self.file_path:
            messagebox.showerror("Error", "Please upload a CSV file first")
            return

        try:
            df = pd.read_csv(self.file_path)
            const.FILE_LOCATIONS['project_csv'] = self.file_path  # Update the path in constants
            const.FILE_LOCATIONS['document_csv'] = self.file_path  # Update the path in constants

            if self.option_vars.get('engineering', tk.BooleanVar(value=False)).get():
                main.engineering_presentation(self.file_path, self.output_folder)
            if self.option_vars.get('impact', tk.BooleanVar(value=False)).get():
                selected_impacted_areas = [area for area, var in self.impacted_areas_vars.items() if var.get()]
                if selected_impacted_areas:
                    main.impact_presentation(self.file_path, selected_impacted_areas, self.output_folder)
                else:
                    main.allimpacted_presentation(self.file_path, self.output_folder)
            if self.option_vars.get('allimpacted', tk.BooleanVar(value=False)).get():
                main.allimpacted_presentation(self.file_path, self.output_folder)
            if self.option_vars.get('who', tk.BooleanVar(value=False)).get():
                main.who_presentation(self.file_path, None, self.output_folder)
            if self.option_vars.get('onhold', tk.BooleanVar(value=False)).get():
                main.onhold_presentation(self.file_path, self.output_folder)
            if self.option_vars.get('objpython ective', tk.BooleanVar(value=False)).get():
                main.objective_presentation(self.file_path, self.output_folder)
            if self.option_vars.get('projects', tk.BooleanVar(value=False)).get():
                main.projects_presentation(self.file_path, self.output_folder)
            if self.option_vars.get('output_all', tk.BooleanVar(value=False)).get():
                main.output_all_presentation(self.file_path, self.output_folder)
            if self.option_vars.get('release', tk.BooleanVar(value=False)).get():
                release_group = self.release_group_var.get()
                if release_group:
                    main.release_presentation(self.file_path, release_group, internal=False, output_folder=self.output_folder)  # Pass the release group as an argument
                else:
                    messagebox.showerror("Error", "Please select a release group for the release report")
            if self.option_vars.get('internal_release', tk.BooleanVar(value=False)).get():
                release_group = self.release_group_var.get()
                if release_group:
                    main.release_presentation(self.file_path, release_group, internal=True, output_folder=self.output_folder)  # Pass the release group and internal flag as arguments
                else:
                    messagebox.showerror("Error", "Please select a release group for the internal release report")
            if self.option_vars.get('docs', tk.BooleanVar(value=False)).get():
                main.docs_presentation(self.file_path, self.output_folder)
            if self.option_vars.get('document_changes', tk.BooleanVar(value=False)).get():
                main.document_changes_presentation(self.file_path, self.output_folder)

        except Exception as e:
            messagebox.showerror("Error", str(e))

def start_gui():
    root = tk.Tk()
    root.title("CSV Processing Tool")
    root.geometry('800x600')  # Set window size
    root.configure(bg='#2c3e50')  # Set background color
    style = ttk.Style(root)
    style.theme_use('clam')
    app = Application(master=root)
    app.mainloop()

if __name__ == "__main__":
    import main
    start_gui()