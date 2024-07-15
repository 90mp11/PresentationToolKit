import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import utilities.constants as const
import utilities.data_utils as du
import utilities.presentation_utils as pu

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.upload_btn = tk.Button(self, text="Upload CSV", command=self.upload_file)
        self.upload_btn.pack(side="top")

        self.quit = tk.Button(self, text="QUIT", fg="red", command=self.master.destroy)
        self.quit.pack(side="bottom")

    def upload_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.process_file(file_path)

    def process_file(self, file_path):
        try:
            df = pd.read_csv(file_path)
            if 'Project Updates' in df.columns:
                self.process_project_file(df)
            elif 'Doc Reference' in df.columns:
                self.process_document_file(df)
            else:
                messagebox.showerror("Error", "Unknown file type")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def process_project_file(self, df):
        # Process the project file
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_ProjectOwner_slides(df, prs)
        output_path = os.path.join(const.FILE_LOCATIONS['output_folder'], "project_presentation.pptx")
        prs.save(output_path)
        messagebox.showinfo("Success", f"Project presentation saved to {output_path}")

    def process_document_file(self, df):
        # Process the document file
        prs = pu.create_blank_presentation(const.FILE_LOCATIONS['pptx_template'])
        pu.create_document_release_section(df, prs)
        output_path = os.path.join(const.FILE_LOCATIONS['output_folder'], "document_presentation.pptx")
        prs.save(output_path)
        messagebox.showinfo("Success", f"Document presentation saved to {output_path}")

root = tk.Tk()
app = Application(master=root)
app.mainloop()
