import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
import os
import datetime
import win32com.client

from change_report import ChangeReport

SMALL_HEIGHT = 130
LARGE_HEIGHT = 160
WIDTH = 600


class ExtractorApp:

    def __init__(self, master):
        self.master = master
        self.folder_path = ""
        self.folder_selected = False
        self.file_path = ""
        self.file_selected = False
        self.errors = []
        self.extension = ""

        self.interface_frame = tk.Frame(self.master)
        self.interface_frame.pack()

        self.file_label = tk.Label(self.interface_frame, text="Select Excel file:")
        self.file_label.grid(row=0, column=0, pady=10)
        self.file_entry = tk.Entry(self.interface_frame, width=50, state="disabled")
        self.file_entry.grid(row=0, column=1)
        self.file_button = tk.Button(self.interface_frame, text="Browse", command=self.select_file)
        self.file_button.grid(row=0, column=2, sticky="ew")

        self.folder_label = tk.Label(self.interface_frame, text="Output folder name:")
        self.folder_label.grid(row=1, column=0, pady=10)
        self.folder_entry = tk.Entry(self.interface_frame, width=50, state="disabled")
        self.folder_entry.grid(row=1, column=1)
        self.folder_button = tk.Button(self.interface_frame, text="Browse", command=self.select_folder)
        self.folder_button.grid(row=1, column=2, sticky="ew")

        self.message_label = tk.Label(self.interface_frame, text="Commit message:")
        self.message_label.grid(row=2, column=0, pady=10)
        self.message_entry = tk.Entry(self.interface_frame, width=50)
        self.message_entry.grid(row=2, column=1)
        self.extract_button = tk.Button(self.interface_frame, text="Extract VBA Code", command=self.extract_code)
        self.extract_button.grid(row=2, column=2)

        # Create a Progressbar widget
        self.progress = ttk.Progressbar(
            self.master, orient=tk.HORIZONTAL, length=200, mode='determinate'
        )

    @property
    def filename(self):
        return self.file_path.split("/")[-1].split(self.extension)[0]

    @property
    def error_string(self):
        message = ""
        if self.errors:
            message += ":\n\n"
            for error in self.errors:
                message += f"{error}\n"

        return message

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm;*.xlam")])
        self.extension = f".{self.file_path.split('.')[-1]}"
        self.file_entry.config(state="normal")
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, self.file_path.split("/")[-1].split(self.extension)[0])
        self.file_selected = True
        self.file_entry.config(state="disabled")

    def select_folder(self):
        self.folder_path = filedialog.askdirectory()
        self.folder_entry.config(state="normal")
        self.folder_entry.delete(0, tk.END)
        self.folder_entry.insert(0, self.folder_path.split("/")[-1])
        self.folder_selected = True
        self.folder_entry.config(state="disabled")

    def extract_code(self):
        if not self.folder_selected:
            messagebox.showwarning("Error", "Please select a file.")
            return

        if not self.file_selected:
            messagebox.showwarning("Error", "Please select a folder.")
            return

        if not self.message_entry.get():
            messagebox.showwarning("Error", "Please enter a commit message.")
            return

        self.progress['value'] = 0
        root.geometry(f"{WIDTH}x{LARGE_HEIGHT}")
        self.progress.pack()
        self.master.update()

        output_dir = self.create_timestamp_folder()
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False

        was_open = self.is_excel_file_open()
        workbook = xl.Workbooks.Open(self.file_path, ReadOnly=True)

        self.progress['maximum'] = len(workbook.VBProject.VBComponents) + 2
        for vbcomp in workbook.VBProject.VBComponents:
            self.progress.step(1)
            self.master.update()
            if not vbcomp.CodeModule.CountOfLines:
                continue

            if vbcomp.Type == 1:  # vbext_ct_StdModule
                extension = ".bas"
                type_folder = "Modules"
            elif vbcomp.Type == 2:  # vbext_ct_ClassModule
                extension = ".cls"
                type_folder = "Class Modules"
            elif vbcomp.Type == 3:  # vbext_ct_MSForm
                extension = ".frm"
                type_folder = "Forms"
            else:
                extension = ".txt"
                type_folder = "Other"

            subfolder_path = os.path.join(output_dir, type_folder)
            if not os.path.exists(subfolder_path):
                os.makedirs(subfolder_path)
            code_file = os.path.join(output_dir, type_folder, f"{vbcomp.Name}{extension}")

            try:
                with open(code_file, "w") as f:
                    for i in range(1, vbcomp.CodeModule.CountOfLines + 1):
                        line = vbcomp.CodeModule.Lines(i, 1)
                        if line.strip() == '':
                            f.write('\n')
                        else:
                            leading_spaces = len(line) - len(line.lstrip())
                            f.write(f"{' ' * leading_spaces}{line.lstrip()}\n")
            except Exception as e:
                self.errors.append(f"Error saving code for module '{vbcomp.Name}': {str(e)}")

        change_report = ChangeReport(self.folder_path)
        change_report.write_change_report()
        self.progress.step(1)
        self.master.update()

        if not was_open:
            workbook.Close(SaveChanges=False)
        xl.Quit()

        self.progress.step(1)
        self.master.update()
        self.progress.pack_forget()
        self.master.update()

        root.geometry(f"{WIDTH}x{SMALL_HEIGHT}")
        messagebox.showinfo(
            "Success",
            f"Completed with {len(self.errors)} error{'' if len(self.errors) == 1 else 's'}{self.error_string}"
        )
        self.message_entry.delete(0, tk.END)
        self.errors = []

    def create_timestamp_folder(self):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H.%M.%S")
        new_folder_path = os.path.join(self.folder_path, f"{timestamp} {self.message_entry.get()}")
        os.makedirs(new_folder_path)
        return new_folder_path

    def is_excel_file_open(self):
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            for wb in excel.Workbooks:
                if wb.FullName == self.file_path:
                    return True
            return False
        except Exception:
            return False


if __name__ == "__main__":
    root = tk.Tk()
    root.title("VBA Code Extractor")
    root.geometry(f"{WIDTH}x{SMALL_HEIGHT}")
    root.resizable(False, False)
    root.lift()

    app_instance = ExtractorApp(root)
    root.mainloop()
