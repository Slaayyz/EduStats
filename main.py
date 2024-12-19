import os
import openpyxl
import matplotlib.pyplot as plt
from openpyxl.utils.exceptions import InvalidFileException

# Checks if the 'Stats' folder exists
def check_directory():
    try:
        if not os.path.exists('Stats'):
            print("The 'Stats' folder does not exist.")
            exit()
        os.chdir('Stats')
    except Exception as e:
        print(f"Error accessing the 'Stats' folder: {e}")
        exit()

# Checks if the necessary files exist or creates 'notes_RT.xlsx'
def check_or_create_files():
    try:
        if not os.path.exists('semestre.xlsx'):
            print("The file 'semestre.xlsx' is not in the 'Stats' folder.")
            exit()

        if not os.path.exists('notes_RT.xlsx'):
            print("The file 'notes_RT.xlsx' is missing. Creating it now...")
            create_notes_file()
        else:
            update_notes_file()
    except Exception as e:
        print(f"Error checking or creating files: {e}")
        exit()

# Creates the 'notes_RT.xlsx' file with a basic structure extracted from 'semestre.xlsx'
def create_notes_file():
    try:
        wb_semestre = openpyxl.load_workbook('semestre.xlsx', data_only=True, read_only=True)
        wb_notes = openpyxl.Workbook()

        for sheet_name in wb_semestre.sheetnames:
            sheet = wb_semestre[sheet_name]
            ws_notes = wb_notes.create_sheet(title=sheet_name)
            ws_notes.append(["Title", "Note1"])

            for row in range(3, sheet.max_row + 1):
                title = sheet.cell(row=row, column=1).value
                if title and isinstance(title, str):
                    ws_notes.append([title])

        if 'Sheet' in wb_notes.sheetnames:
            del wb_notes['Sheet']

        wb_notes.save('notes_RT.xlsx')
        wb_notes.close()
        wb_semestre.close()
        print("The file 'notes_RT.xlsx' was successfully created.")
    except InvalidFileException:
        print("Error: 'semestre.xlsx' is corrupted or invalid.")
        exit()
    except Exception as e:
        print(f"Error creating 'notes_RT.xlsx': {e}")
        exit()

# Updates the 'notes_RT.xlsx' file if changes were made to 'semestre.xlsx'
def update_notes_file():
    try:
        wb_semestre = openpyxl.load_workbook('semestre.xlsx', data_only=True, read_only=True)
        wb_notes = openpyxl.load_workbook('notes_RT.xlsx')

        for sheet_name in wb_semestre.sheetnames:
            sheet = wb_semestre[sheet_name]
            if sheet_name not in wb_notes.sheetnames:
                ws_notes = wb_notes.create_sheet(title=sheet_name)
                ws_notes.append(["Title", "Note1", "Note2", "Note3"])
            else:
                ws_notes = wb_notes[sheet_name]

            existing_titles = {ws_notes.cell(row=row, column=1).value for row in range(2, ws_notes.max_row + 1)}

            for row in range(3, sheet.max_row + 1):
                title = sheet.cell(row=row, column=1).value
                if title and isinstance(title, str) and title not in existing_titles:
                    ws_notes.append([title])

        wb_notes.save('notes_RT.xlsx')
        wb_notes.close()
        wb_semestre.close()
        print("The file 'notes_RT.xlsx' was successfully updated.")
    except InvalidFileException:
        print("Error: 'semestre.xlsx' or 'notes_RT.xlsx' is corrupted or invalid.")
        exit()
    except Exception as e:
        print(f"Error updating 'notes_RT.xlsx': {e}")
        exit()

# Loads the Excel files at once and reads all data
def load_workbooks():
    try:
        wb_file = openpyxl.load_workbook('semestre.xlsx', data_only=True, read_only=True)
        wb_note = openpyxl.load_workbook('notes_RT.xlsx', data_only=True, read_only=True)
        return wb_file, wb_note
    except InvalidFileException:
        print("Error: One of the Excel files is corrupted or invalid.")
        exit()
    except Exception as e:
        print(f"Error loading the Excel files: {e}")
        exit()

# Calculates the average per UE (optimized)
def calculate_averages(wb_file, wb_note):
    try:
        UE_S = []
        notes_dict = {}

        for sheet_name in wb_note.sheetnames:
            sheet_notes = wb_note[sheet_name]  # Retrieve the corresponding notes sheet

            # Prepare a dictionary of notes by Title
            for row in range(2, sheet_notes.max_row + 1):  # Skip the header row
                title = str(sheet_notes.cell(row=row, column=1).value)
                notes_dict[title] = []
                for col in range(2, sheet_notes.max_column + 1):
                    note = sheet_notes.cell(row=row, column=col).value
                    if isinstance(note, (int, float)):  # Keep only numerical notes
                        notes_dict[title].append(note)

        for sheet_name in wb_file.sheetnames:
            sheet_file = wb_file[sheet_name]
            for col in range(1, sheet_file.max_column + 1):
                for row in range(1, sheet_file.max_row + 1):
                    cell = sheet_file.cell(row=row, column=col)
                    if cell.value and cell.data_type == 's' and cell.value.startswith('UE'):
                        UE = cell.value
                        UEden, UEnom = 0, 0
                        # Read column data once to avoid multiple iterations
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
        print(f"Error calculating averages: {e}")
        return []

# Function to display the graph
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
        print(f"Error generating the graph: {e}")

# Main function
def main():
    print("Current directory:", os.getcwd())
    check_directory()
    check_or_create_files()

    wb_file, wb_note = load_workbooks()
    UE_averages = calculate_averages(wb_file, wb_note)
    if UE_averages:
        plot_results(UE_averages)
    else:
        print("No averages calculated. Check your files.")

    wb_file.close()
    wb_note.close()

if __name__ == "__main__":
    main()
