import os
import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
import logging
from openpyxl.utils.exceptions import InvalidFileException

# Constants
STATS_FOLDER = 'stats'
SEMESTER_FILE = 'semester.xlsx'
NOTES_FILE = 'notes_RT.xlsx'

# Logger setup
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def check_and_change_directory():
    if not os.path.exists(STATS_FOLDER):
        os.makedirs(STATS_FOLDER)
    os.chdir(STATS_FOLDER)

def validate_notes_file():
    if not os.path.exists(NOTES_FILE):
        messagebox.showerror("Error", f"The file '{NOTES_FILE}' is missing.")
        exit()

def save_changes(tree, sheet_name):
    wb_notes = openpyxl.load_workbook(NOTES_FILE)
    ws_notes = wb_notes[sheet_name]

    for row in ws_notes.iter_rows(min_row=2, max_row=ws_notes.max_row, max_col=2):
        for cell in row:
            cell.value = None

    for i, item in enumerate(tree.get_children()):
        values = tree.item(item, 'values')
        title = values[0]
        note = float(values[1]) if values[1] else None
        ws_notes.cell(row=i + 2, column=1).value = title  # Subject
        ws_notes.cell(row=i + 2, column=2).value = note  # Note

    wb_notes.save(NOTES_FILE)
    wb_notes.close()

def run_secondary_script():
    def load_workbook(file_name, read_only=True):
        try:
            return openpyxl.load_workbook(file_name, data_only=True, read_only=read_only)
        except InvalidFileException:
            logging.error(f"Error: '{file_name}' is corrupted or invalid.")
            exit()
        except Exception as e:
            logging.error(f"Error loading the file '{file_name}': {e}")
            exit()

    def validate_files():
        if not os.path.exists(SEMESTER_FILE):
            logging.error(f"The file '{SEMESTER_FILE}' is not in the '{STATS_FOLDER}' folder.")
            exit()

    def calculate_averages(wb_file, wb_note):
        try:
            UE_S = []
            notes_dict = {}

            for sheet_name in wb_note.sheetnames:
                sheet_notes = wb_note[sheet_name]

                for row in range(2, sheet_notes.max_row + 1):
                    title = str(sheet_notes.cell(row=row, column=1).value)
                    notes_dict[title] = []
                    for col in range(2, sheet_notes.max_column + 1):
                        note = sheet_notes.cell(row=row, column=col).value
                        if isinstance(note, (int, float)):
                            notes_dict[title].append(note)

            for sheet_name in wb_file.sheetnames:
                sheet_file = wb_file[sheet_name]

                for col in range(1, sheet_file.max_column + 1):
                    for row in range(1, sheet_file.max_row + 1):
                        cell = sheet_file.cell(row=row, column=col)
                        if cell.value and cell.data_type == 's' and cell.value.startswith('UE'):
                            UE = cell.value
                            UEden, UEnom = 0, 0

                            for i in range(3, sheet_file.max_row + 1):
                                coefficient = sheet_file.cell(row=i, column=col).value
                                title = str(sheet_file.cell(row=i, column=1).value)
                                if coefficient and title and title in notes_dict:
                                    for note in notes_dict[title]:
                                        UEnom += note * coefficient
                                        UEden += coefficient

                            if UEden != 0:
                                UE_S.append((UE, UEnom / UEden))

            return UE_S
        except Exception as e:
            logging.error(f"Error calculating averages: {e}")
            return []

    def plot_results(values):
        try:
            plt.figure()
            plt.title("Results of BUT RT")
            plt.ylim(0, 20)
            plt.axhline(10, color="Red", linestyle='--', label='Pass Threshold')
            plt.axhline(8, color="Orange", linestyle='--', label='Retake Threshold')

            for UE, avg in values:
                color = "Green" if avg >= 10 else "Orange" if avg >= 8 else "Red"
                plt.bar(UE, avg, color=color)

            plt.legend()
            plt.show()
        except Exception as e:
            logging.error(f"Error generating the graph: {e}")

    validate_files()

    wb_file = load_workbook(SEMESTER_FILE)
    wb_note = load_workbook(NOTES_FILE)

    UE_averages = calculate_averages(wb_file, wb_note)
    if UE_averages:
        plot_results(UE_averages)
    else:
        logging.warning("No averages calculated. Check your files.")

    wb_file.close()
    wb_note.close()

def display_gui():
    wb_notes = openpyxl.load_workbook(NOTES_FILE)

    root = tk.Tk()
    root.title("Note Management")

    notebook = ttk.Notebook(root)
    notebook.pack(fill='both', expand=True)

    for sheet_name in wb_notes.sheetnames:
        sheet = wb_notes[sheet_name]

        frame = ttk.Frame(notebook)
        notebook.add(frame, text=sheet_name)

        tree = ttk.Treeview(frame, columns=("Title", "Note"), show='headings')
        tree.heading("Title", text="Subject")
        tree.heading("Note", text="Note")
        tree.pack(fill='both', expand=True)

        for row in sheet.iter_rows(min_row=2, values_only=True):
            tree.insert("", "end", values=row)

        def save_sheet(sheet_name=sheet_name, tree=tree):
            save_changes(tree, sheet_name)
            messagebox.showinfo("Success", f"Changes have been saved for {sheet_name}.")

        save_button = ttk.Button(frame, text="Save Changes", command=save_sheet)
        save_button.pack()

        def edit_selected_item(event, tree=tree):
            selected_item = tree.focus()
            if selected_item:
                values = tree.item(selected_item, 'values')
                title = values[0]
                note = values[1]

                edit_window = tk.Toplevel(root)
                edit_window.title("Edit Note")

                tk.Label(edit_window, text=f"Subject: {title}").pack(pady=10)
                tk.Label(edit_window, text="New Note:").pack(pady=5)
                note_entry = tk.Entry(edit_window)
                note_entry.insert(0, note if note else "")
                note_entry.pack(pady=5)

                def save_edit():
                    new_note = note_entry.get()
                    tree.item(selected_item, values=(title, new_note))
                    edit_window.destroy()

                tk.Button(edit_window, text="Save", command=save_edit).pack(pady=10)

        tree.bind("<Double-1>", edit_selected_item)

    run_button = ttk.Button(root, text="Run Note Analysis", command=run_secondary_script)
    run_button.pack(pady=10)

    root.mainloop()

def main():
    check_and_change_directory()
    validate_notes_file()
    display_gui()

if __name__ == "__main__":
    main()
