import os
import sys

import openpyxl
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QMessageBox, QLabel, QListWidget, QPushButton, QGridLayout, QFileDialog, QTextEdit
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize
import pandas as pd
import csv
import re
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.styles.borders import Border, Side
from openpyxl.utils.cell import get_column_letter
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
import chardet

COMMENT_CODE = 4444153
# STATEMENT_CODE = 4389991
STATEMENT_CODE = 4442358
COURSE_NUMBER = 110

class MainWindow(QMainWindow):
    def __init__(self):
        # Ensures all initialization code from the inherited class is executed
        super().__init__()

        # Global Variables
        self.data = pd.DataFrame()
        self.final_data = pd.DataFrame()
        self.roster = pd.DataFrame()
        self.instructors = None
        self.assignment = None
        self.points = None
        self.questions = None
        self.hasComments = False
        self.hasRoster = False
        self.hasInstructorData = False
        self.hasSectionData = False
        self.colComments = None
        self.colStatement = None
        self.sections = None
        self.output_dir = None
        self.assignment_output_dir = None
        self.output_wb = None

        # Setup output directory
        self.output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        # Initialize main widget
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.setWindowTitle("Cengage Data Scraper")

        # Create the layout
        main_layout = QGridLayout()

        # Labels
        self.lblRoster = QLabel("Current Roster")
        self.lblSections = QLabel("Sections")
        self.lblInstructors = QLabel("Instructors")
        self.lblData = QLabel("Data")

        self.lblRoster.setFixedHeight(10)
        self.lblSections.setFixedHeight(10)
        self.lblInstructors.setFixedHeight(10)
        self.lblData.setFixedHeight(10)

        # Lists and one text box
        self.listRoster = QListWidget()
        self.listSections = QListWidget()
        self.listInstructors = QListWidget()
        self.txtData = QTextEdit()

        self.listSections.setFixedHeight(200)
        self.listInstructors.setFixedHeight(200)
        self.txtData.setFixedWidth(500)
        self.txtData.setLineWrapMode(QTextEdit.NoWrap)

        self.listRoster.setStyleSheet("border: 1px solid grey;")
        self.listSections.setStyleSheet("border: 1px solid grey;")
        self.listInstructors.setStyleSheet("border: 1px solid grey;")
        self.txtData.setStyleSheet("border: 1px solid grey;")

        # Buttons
        self.btnRoster = QPushButton("Load Roster")
        self.btnSections = QPushButton("Assign Sections")
        self.btnRemoveInstructors = QPushButton("Remove Instructor")
        self.btnLoadData = QPushButton("Load Data")
        self.btnExportData = QPushButton("Export Data")

        self.btnRoster.setEnabled(True)
        self.btnSections.setEnabled(False)
        self.btnRemoveInstructors.setEnabled(False)
        self.btnLoadData.setEnabled(False)
        self.btnExportData.setEnabled(False)

        btnMargin = 25
        btnIconHeight = 32
        btnIconSize = QSize(btnIconHeight, btnIconHeight)
        btnHeight = btnIconHeight + btnMargin

        self.btnRoster.setFixedHeight(btnHeight)
        self.btnSections.setFixedHeight(btnHeight)
        self.btnRemoveInstructors.setFixedHeight(btnHeight)
        self.btnLoadData.setFixedHeight(btnHeight)
        self.btnExportData.setFixedHeight(btnHeight)

        self.btnLoadData.setIcon(QIcon('resources/icons/parse.ico'))
        self.btnLoadData.setIconSize(btnIconSize)
        self.btnExportData.setIcon(QIcon('resources/icons/export.png'))
        self.btnExportData.setIconSize(btnIconSize)

        # Connect buttons to functions
        self.btnRoster.clicked.connect(self.setup_roster)
        self.btnSections.clicked.connect(self.assign_sections)
        self.btnRemoveInstructors.clicked.connect(self.remove_instructors)
        self.btnLoadData.clicked.connect(self.load_file)
        self.btnExportData.clicked.connect(self.export)

        # Add widgets to the grid
        main_layout.addWidget(self.lblRoster, 0, 0)
        main_layout.addWidget(self.lblSections, 0, 1)
        main_layout.addWidget(self.lblData, 0, 3)

        main_layout.addWidget(self.listRoster, 1, 0, 3, 1)
        main_layout.addWidget(self.listSections, 1, 1, 1, 2)
        main_layout.addWidget(self.lblInstructors, 2, 1)
        main_layout.addWidget(self.listInstructors, 3, 1, 1, 2)
        main_layout.addWidget(self.txtData, 1, 3, 3, 2)

        main_layout.addWidget(self.btnRoster, 4, 0)
        main_layout.addWidget(self.btnRemoveInstructors, 4, 1)
        main_layout.addWidget(self.btnSections, 4, 2)
        main_layout.addWidget(self.btnLoadData, 4, 3)
        main_layout.addWidget(self.btnExportData, 4, 4)

        # Set the layout for the GUI
        main_widget.setLayout(main_layout)

        # Check for configuration files (roster, instructors, and sections)
        self.check_for_config_files()

    def show_message(self, message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(message)
        msg.setWindowTitle("Message")
        msg.exec_()

    def show_error(self, message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(message)
        msg.setWindowTitle("Error")
        msg.exec_()

    def confirm_action(self, message):
        reply = QMessageBox.question(self, 'Confirmation', message, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            return True
        else:
            return False

    def open_file_dialog(self, extensions):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File", "", extensions, options=options)

        if file_name:
            return file_name

    def strip_whitespace(self, x):
        if isinstance(x, str):
            return x.strip()
        return x

    def process_names(self, text):
        # Use a regular expression to find the pattern and preserve the desired part
        match = re.search(r',[^ ]+', text)
        if match:
            end_index = match.end()
            first_space_after_comma = text.find(' ', end_index)
            if first_space_after_comma != -1:
                text = text[:first_space_after_comma]
            else:
                text = text[:end_index]
        return text

    def load_file(self):
        # Check if roster loaded, exit function if no roster
        if self.listRoster.count() == 0:
            self.show_message(f"No roster found, please generate a new roster.")
            return

        # Open file dialog
        file_path = self.open_file_dialog("CSV files (*.csv)")

        # Comment out if the encoding causes issues
        with open(file_path,'rb') as f:
            result = chardet.detect(f.read())
            encoding = result['encoding']

        if file_path:
            try:
                header_lines = []
                if file_path.endswith('.csv'):
                    with open(file_path, 'r') as file:
                        for _ in range(8):
                            header_lines.append(file.readline().strip())
                    # Process the header information
                    self.process_header(header_lines)

                    # Load the remaining file into a DataFrame, skipping the first 9 lines
                    df = pd.read_csv(file_path, skiprows=9, header=None, encoding=encoding)

                    # Use if the encoding causes issues
                    # df = pd.read_csv(file_path, skiprows=9, header=None)

                else:
                    raise ValueError("Unsupported file format")

                # Display DataFrame in a new window
                self.data = df
                self.parse_data()

                # Enable export button
                self.btnExportData.setEnabled(True)

            except Exception as e:
                self.show_error(f"Failed to load file\n{e}\n{file_path}")
                return

    def process_header(self, header_lines):
        # Use csv.reader to properly handle commas within quotes
        header_info = [list(csv.reader([line]))[0] for line in header_lines]
        self.assignment = header_info[4][1]
        self.points = [float(x) for x in header_info[7] if x != '' and x != 'Points']
        self.questions = len(self.points)

        j = 0
        for i in header_info[6]:
            if i == f'{COMMENT_CODE}':
                self.hasComments = True
                self.colComments = j
                self.questions -= 1
                j += 1
            elif i == f'{STATEMENT_CODE}':
                self.colStatement = j
                self.questions -= 1
                j += 1
            else:
                j += 1

        # print(self.hasComments)
        # print(self.colComments)
        # print(self.assignment)
        # print(self.points)

    def parse_data(self):
        if self.data.empty: return

        df_copy = self.data

        # Iterate over every two lines
        for index in range(0, len(df_copy), 2):
            row1 = df_copy.iloc[index]
            if index + 1 < len(df_copy):
                row2 = df_copy.iloc[index + 1]

                # Save the total score
                df_copy.iloc[index, 3] = row2[3]

                for i in range(4, len(row2)-1,1):
                    if (self.hasComments and i == self.colComments) or i == self.colStatement:
                        continue
                    if float(row2[i]) == 0 and str(row1[i]) != "nan":
                        df_copy.iloc[index, i] = "0"
                    elif float(row2[i]) == 0 and str(row1[i]) == "nan":
                        df_copy.iloc[index, i] = "-"
                    else:
                        df_copy.iloc[index, i] = "1"
            else:
                print("Unpaired row...")
        
        filter_second_row = df_copy.iloc[:,0].notna()
        df_filtered = df_copy[filter_second_row]
        df_filtered.loc[:, 1] = df_filtered[1].str.replace('@usafa$', '', regex=True)
        df_filtered.loc[:, 0] = df_filtered[0].str.replace(', ', ',', regex=False)
        df_filtered.loc[:, 0] = df_filtered[0].map(self.process_names)
        df_filtered = df_filtered.fillna('')
        # df_filtered.loc[:, 2] = None
        # df_filtered.loc[:, 2] = ""
        # df_filtered = df_filtered.drop(3, axis=1)
        
        # NEW CODE: Reorder columns to put comments and statement at end
        cols = list(df_filtered.columns)

        # Create new column order without the comments and statement columns
        new_cols = [col for col in cols if col not in [self.colComments, self.colStatement]]
        print(new_cols)
        print(f'Comments: {self.colComments}, Statement: {self.colStatement}')
        # Add comments and statement columns back in desired position
        if self.hasComments:
            new_cols.insert(len(new_cols) - 1, self.colComments)  # Second to last

        new_cols.insert(len(new_cols) - 1, self.colStatement)     # Last before final column

        print(new_cols)
        # Reorder the DataFrame
        df_filtered = df_filtered[new_cols]

        # Renumber the columns sequentially
        df_filtered.columns = list(range(len(new_cols)))

        self.final_data = df_filtered

        email_filter = self.final_data.iloc[:, 1].isin(self.roster.iloc[:, 1])
        name_filter = self.final_data.iloc[:, 0].isin(self.roster.iloc[:, 2])
        combined_filter = email_filter | name_filter

        if self.hasInstructorData == False:
            self.instructors = self.final_data.iloc[:, 0:2][~combined_filter]
            self.save_instructors()

        self.final_data = self.final_data[combined_filter]

        for index, row in self.final_data.iterrows():

            filter_one = self.roster["Email"] == row[1]
            if filter_one.any():
                self.final_data.at[index, 2] = self.roster.loc[filter_one, "Section"].values[0]
                continue

            filter_two = self.roster["Cadet Name"] == row[0]
            if filter_two.any():
                self.final_data.at[index, 2] = self.roster.loc[filter_two, "Section"].values[0]
                continue

        # Display in text edit
        df_display = self.final_data.drop(1, axis=1)
        content = df_display.to_string(index=False, header=False)
        self.txtData.setPlainText(content)

    def setup_roster(self):
        # Open file dialog
        file_path = self.open_file_dialog("Excel Files (*.xlsx *.xls)")

        if file_path:
            try:
                df = pd.read_excel(file_path, skiprows=1)
                print(df)
                df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
                filter = df["Course Number"] == f'{COURSE_NUMBER}'
                print(df[filter])
                df_filtered = df[filter]
                df_filtered = df_filtered[["Section", "Email", "Cadet Name"]]
                df_filtered["Cadet Name"] = df_filtered["Cadet Name"].map(self.process_names)
                print(df_filtered)
                # Get the script directory
                script_dir = os.path.dirname(os.path.abspath(__file__))

                # Define the file path in the script directory
                file_path = os.path.join(script_dir, 'current_roster.csv')

                # Save the DataFrame to a CSV file
                df_filtered.to_csv(file_path, index=False)

                print(f"Roster info saved to {file_path}")

                self.check_for_roster()
                self.check_for_sections()

            except Exception as e:
                self.show_error(f"Failed to load file\n{e}")

    def save_instructors(self):
        # Get the script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # Define the file path in the script directory
        file_path = os.path.join(script_dir, 'instructors.csv')

        # Save the DataFrame to a CSV file
        self.instructors.to_csv(file_path, index=False)

        print(f"Instructors info saved to {file_path}")
        self.check_for_instructors()

    def save_sections(self):
        # Get the script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # Define the file path in the script directory
        file_path = os.path.join(script_dir, 'sections.csv')

        # Save the DataFrame to a CSV file
        self.sections.to_csv(file_path, index=False)

        print(f"Sections info saved to {file_path}")
        self.check_for_sections()

    def check_for_config_files(self):
        self.check_for_roster()
        self.check_for_instructors()
        self.check_for_sections()

    def check_for_roster(self):
        # Get the script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_dir, 'current_roster.csv')

        # Check if the file exists
        if os.path.exists(file_path):
            # Load the file into a DataFrame
            df = pd.read_csv(file_path)

            # Save to roster global
            self.roster = df
            self.sections = self.roster.iloc[:, 0].unique()

            # Clear the Listbox
            self.listRoster.clear()
            # self.listSections.clear()

            # Insert DataFrame content into the Listbox
            for index, row in self.roster.iterrows():
                self.listRoster.addItem(f"{row['Cadet Name']} ({row['Section']})")

            # # Insert DataFrame content into the Listbox
            # for row in self.sections:
            #     self.listSections.addItem(row)

            self.btnRoster.setText("Update Roster")

            # Enable appropriate buttons after load
            self.btnSections.setEnabled(True)
            self.btnLoadData.setEnabled(True)
            self.hasRoster = True
        else:
            self.show_message(f"No roster found, please generate a new roster.")

    def check_for_instructors(self):
        # Get the script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_dir, 'instructors.csv')

        # Check if the file exists
        if os.path.exists(file_path):
            # Load the file into a DataFrame
            df = pd.read_csv(file_path)

            # Save to roster global
            self.instructors = df

            # Clear the Listbox
            self.listInstructors.clear()

            # Insert DataFrame content into the Listbox
            for index, row in self.instructors.iterrows():
                self.listInstructors.addItem(f"{row.iloc[0]}")

            # Enable appropriate buttons after load
            self.btnSections.setEnabled(True)
            self.btnRemoveInstructors.setEnabled(True)
            self.hasInstructorData = True
        else:
            self.show_message(f"No instructors found, please load a file to generate.")

    def check_for_sections(self):
        # Get the script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(script_dir, 'sections.csv')

        # Check if the file exists
        if os.path.exists(file_path):
            # Load the file into a DataFrame
            df = pd.read_csv(file_path)

            # Save to roster global
            self.sections = df

            # Clear the Listbox
            self.listSections.clear()

            # Insert DataFrame content into the Listbox
            for index, row in self.sections.iterrows():
                if row.iloc[1] == " ":
                    self.listSections.addItem(f"{row.iloc[0]}")
                else:
                    self.listSections.addItem(f"{row.iloc[0]} -- {row.iloc[1]}")

            # Enable appropriate buttons after load
            self.btnSections.setEnabled(True)
            self.hasSectionData = True
        elif self.hasRoster:
            self.sections = self.roster.iloc[:, 0].unique()
            self.sections = pd.DataFrame(self.sections)
            self.sections[1] = str(" ")
            self.save_sections()
        else:
            self.show_message(f"No section info found, please reload roster to generate.")

    def remove_instructors(self):
        selected = self.listInstructors.currentItem().text()

        if not self.confirm_action(f'Are you sure you want to remove {selected} from the list of instructors?'):
            return

        self.instructors = self.instructors[~(self.instructors.iloc[:, 0] == selected)]
        self.save_instructors()

    def assign_sections(self):
        j = 0
        if self.listSections.selectedItems() and self.listInstructors.selectedItems():
            filter = self.sections.iloc[:, 0] == self.listSections.currentItem().text().split(" -- ")[0]            # Have to split on ' -- ' in case the section has already been assigned
            index = self.sections[filter].index[0]
            self.sections.iloc[index, 1] = self.listInstructors.currentItem().text()
            self.save_sections()
        else:
            self.show_error("Must select a section and an instructor")

    def pixel_to_pt(self, x):
        return x / 7.0

    def truncate_string(self, s, max_length=20):
        return s if len(s) <= max_length else s[:max_length] + '...'

    def truncate_or_pad_string(self, s, max_length=70):
        return (s[:max_length - 3] + '...') if len(s) > max_length else s.ljust(max_length)

    def export(self):
        # Create path to assignment output folder and create if it doesn't exist
        self.assignment_output_dir = os.path.join(self.output_dir, f'{self.assignment.split("/")[0].strip()}')
        print(self.assignment_output_dir)
        result = f'Error saving files'

        if not os.path.exists(self.assignment_output_dir):
            os.makedirs(self.assignment_output_dir)

        for section in self.sections.iloc[:, 0]:
            subset = self.final_data[self.final_data.iloc[:, 2] == section]

            if not subset.empty:
                result = self.create_table(subset)

        print(f'Saved the output to {result}')


    def create_pdf(self, ws, title, maxrow):
        # Extract data and cell colors from the worksheet, ignoring the first row, last row, and first/last columns

        data = []
        cell_colors = []
        for row in ws.iter_rows(min_row=2, max_row=maxrow, min_col=2, max_col=ws.max_column - 1):
            row_data = []
            row_colors = []
            for cell in row:
                row_data.append(str(cell.value).lstrip() if cell.value is not None else "")
                if cell.fill.fgColor.rgb is not None:
                    # Convert RGB from 'FF123456' to hex '#123456'
                    hex_color = '#' + cell.fill.fgColor.rgb[2:]
                else:
                    hex_color = None
                row_colors.append(hex_color)
            data.append(row_data)
            cell_colors.append(row_colors)

        # Create a PDF document
        pdf_file = f'{title}.pdf'
        file_path = os.path.join(self.assignment_output_dir, pdf_file)
        # pdf = SimpleDocTemplate(pdf_file, pagesize=landscape(letter))
        pdf = SimpleDocTemplate(file_path, pagesize=landscape(letter), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)

        # Create a table with the extracted data
        table = Table(data)

        # Define basic table style
        style = TableStyle([
            ('GRID', (0, 2), (-1, -1), 1, colors.black),  # Grid color and thickness
            ('BOX', (0, 0), (-1, -1), 2, colors.black),  # Outer box color and thickness
        ])

        # Apply style to the table
        table.setStyle(style)

        # Apply specific alignments and colors
        for row_idx, row in enumerate(data):
            for col_idx, _ in enumerate(row):
                # Alignment
                if col_idx == 0 or col_idx >= len(row) - 2:  # First column or last two columns
                    alignment = 'LEFT'
                    if row_idx == 0 and col_idx == len(row) - 2:  # First row, second to last column
                        alignment = 'RIGHT'
                else:
                    alignment = 'CENTER'
                table.setStyle(TableStyle([
                    ('ALIGN', (col_idx, row_idx), (col_idx, row_idx), alignment)
                ]))
                # Background color
                if cell_colors[row_idx][col_idx] is not None:
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (col_idx, row_idx), (col_idx, row_idx), cell_colors[row_idx][col_idx])
                    ]))

        # Build the PDF
        elements = [table]
        pdf.build(elements)

        print(f'PDF created successfully. [{file_path}]')

    def create_table(self, dfOutput):
        # Define the file path in the script directory
        file_path = os.path.join(self.assignment_output_dir, f'output.xlsx')

        if os.path.exists(file_path):
            self.output_wb = openpyxl.load_workbook(file_path)
        else:
            self.output_wb = openpyxl.Workbook()

        currentSection =self.sections[self.sections.iloc[:, 0] == dfOutput.iloc[0, 2]]

        # Create a new worksheet
        if self.output_wb.sheetnames[0] == 'Sheet':
            ws = self.output_wb.active
            ws.title = f'{currentSection.iloc[0,0]}'
        elif currentSection.iloc[0,0] in self.output_wb.sheetnames:
            ws = self.output_wb[f'{currentSection.iloc[0,0]}']
        else:
            ws = self.output_wb.create_sheet(title=f'{currentSection.iloc[0,0]}')

        dfOutput = dfOutput.drop(1, axis=1)
        dfOutput = dfOutput.drop(2, axis=1)

        # define fill colors
        greenFill = PatternFill(start_color='FF00B050', end_color='FF00B050', fill_type='solid')
        redFill = PatternFill(start_color='FFC00000', end_color='FFC00000', fill_type='solid')
        borderFill = PatternFill(start_color='FF808080', end_color='FF808080', fill_type='solid')
        headerFill = PatternFill(start_color='FFD9D9D9', end_color='FFD9D9D9', fill_type='solid')
        whiteFill = PatternFill(start_color='FFFFFFFF', end_color='FFFFFFFF', fill_type='solid')
        warningFill = PatternFill(start_color='FFF59412', end_color='FFF59412', fill_type='solid')

        # Define border styles
        thin_all_sides = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        thin_header_left = Border(left=Side(style='thin'), bottom=Side(style='thin'))
        thin_header_middle = Border(bottom=Side(style='thin'))
        thin_header_right = Border(right=Side(style='thin'), bottom=Side(style='thin'))
        thin_header_left_top = Border(left=Side(style='thin'), top=Side(style='thin'))
        thin_header_middle_top = Border(top=Side(style='thin'))
        thin_header_right_top = Border(right=Side(style='thin'), top=Side(style='thin'))

        # Specify standard geometry
        default_row_height = 18
        number_of_columns = len(dfOutput.columns)
        number_of_rows = len(dfOutput.index)
        headerRows = 2
        offsetRows = 3
        commentCol = None

        if self.hasComments:
            offsetCols = 3
        else:
            offsetCols = 2

        # Specify the row geometry
        for i in range(1, 201, 1):
            ws.row_dimensions[i].height = default_row_height

        # Specify column geometry
        ws.column_dimensions[get_column_letter(1)].width = self.pixel_to_pt(21)
        ws.column_dimensions[get_column_letter(2)].width = self.pixel_to_pt(168)
        ws.column_dimensions[get_column_letter(3)].width = self.pixel_to_pt(42)

        i = 0
        for i in range(4, number_of_columns - 1):
            ws.column_dimensions[get_column_letter(i)].width = self.pixel_to_pt(19)

        if self.hasComments:
            ws.column_dimensions[get_column_letter(i + 1)].width = self.pixel_to_pt(665)
            ws.column_dimensions[get_column_letter(i + 2)].width = self.pixel_to_pt(294)
            ws.column_dimensions[get_column_letter(i + 3)].width = self.pixel_to_pt(21)

            # If statement to check for hasComments boolean
            lastcolumn = i + 3

        else:
            ws.column_dimensions[get_column_letter(i + 1)].width = self.pixel_to_pt(294)
            ws.column_dimensions[get_column_letter(i + 2)].width = self.pixel_to_pt(21)

            # If statement to check for hasComments boolean
            lastcolumn = i + 2

        # Set fill colors
        # Borders
        for col in range(1, lastcolumn + 1):
            for row in range(1, number_of_rows + headerRows + offsetRows):
                ws[get_column_letter(col) + str(row)].fill = borderFill

        # Interior
        for col in range(2, lastcolumn):
            for row in range(2, number_of_rows + headerRows + offsetRows - 1):
                index = get_column_letter(col) + str(row)
                if row == 2:
                    ws[index].alignment = Alignment(vertical='center')
                    ws[index].font = Font(bold=True)
                    ws[index].fill = headerFill
                    if col == 2:
                        ws[index] = f' {currentSection.iloc[0,0]} -- {currentSection.iloc[0,1].split(",")[0]}'
                        ws[index].border = thin_header_left_top
                    elif col == lastcolumn - 1:
                        ws[index].border = thin_header_right_top
                    else:
                        ws[index].border = thin_header_middle_top

                        if col == 3:
                            pass
                        elif col >= 4 and col <= number_of_columns + 1 - offsetCols:
                            pass
                        else:
                            ws[index] = f'{self.assignment}'
                            ws[index].alignment = Alignment(horizontal='center', vertical='center')
                elif row == 3:
                    ws[index].alignment = Alignment(vertical='center')
                    ws[index].font = Font(bold=True)
                    ws[index].fill = headerFill
                    if col == 2:
                        ws[index].border = thin_header_left
                        ws[index] = " Name"
                    elif col == lastcolumn - 1:
                        ws[index].border = thin_header_right
                        ws[index] = " Documentation"
                    else:
                        ws[index].border = thin_header_middle

                        if col == 3:
                            ws[index] = "Total"
                            ws[index].alignment = Alignment(horizontal='center', vertical='center')
                        elif col >= 4 and col <= number_of_columns + 1 - offsetCols:
                            ws[index] = f'Q{col - offsetCols}'
                            ws[index].alignment = Alignment(horizontal='center', vertical='center')
                        else:
                            ws[index] = "  What did you find interesting/useful/confusing?"
                            commentCol = col

                elif row - 4 >= 0 and row - 4 < number_of_rows:
                    df_row = row - 4
                    df_col = col - 2
                    data = dfOutput.iloc[df_row, df_col]
                    ws[index].alignment = Alignment(vertical='center')

                    if df_col == 0:
                        ws[index] = data.replace(',', ', ')
                        ws[index].fill = whiteFill
                    elif df_col == 1:
                        ws[index] = data
                        ws[index].fill = whiteFill
                        ws[index].alignment = Alignment(horizontal='center', vertical='center')
                    elif df_col >= 2 and df_col < number_of_columns - offsetCols:
                        ws[index] = None
                        if data == '1' :
                            ws[index].fill = greenFill
                        elif data == '0':
                            ws[index].fill = redFill
                        elif data == '-':
                            ws[index] = data
                            ws[index].fill = whiteFill
                            ws[index].alignment = Alignment(horizontal='center', vertical='center')
                        else:
                            ws[index] = data
                            ws[index].fill = warningFill
                    else:
                        if df_col == number_of_columns - offsetCols:
                            max_length = 70
                        else:
                            max_length = 40
                        ws[index] = self.truncate_or_pad_string(data, max_length=max_length)
                        ws[index].fill = whiteFill

                    ws[index].border = thin_all_sides
                else:
                    ws[index].fill = whiteFill

        # Generate PDF from current worksheet
        # self.create_pdf(ws, currentSection.iloc[0, 0], number_of_rows + 3)

        # Add the copy-paste section to the worksheet
        if self.hasComments:
            comments_to_filter = ['', ' ', 'no', 'nothing yet', 'none', 'nothing', 'unsure']
            n0 = number_of_rows + 8
            n = n0
            index = get_column_letter(commentCol) + str(n)
            ws[index] = f'Copy and Paste Comments:'

            for comment in dfOutput.iloc[:, number_of_columns - 3]:
                if comment.lower() in (item.lower() for item in comments_to_filter):
                    continue
                n += 1
                index = get_column_letter(commentCol) + str(n)
                ws[index] = self.truncate_string(comment, max_length=500)
                # ws[index].alignment = Alignment(wrap_text=True)

            # Apply border to the region
            for row in range(n0 + 1, n + 1):
                for col in range(commentCol, commentCol + 1):
                    cell = ws.cell(row=row, column=col)
                    if row == n0 + 1:
                        cell.border = Border(top=thin_all_sides.top, left=thin_all_sides.left, right=thin_all_sides.right)
                    elif row == n:
                        cell.border = Border(bottom=thin_all_sides.bottom, left=thin_all_sides.left, right=thin_all_sides.right)
                    else:
                        cell.border = Border(left=thin_all_sides.left, right=thin_all_sides.right)

        # Save the workbook
        self.output_wb.save(file_path)

        return file_path

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())