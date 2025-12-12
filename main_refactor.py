import os
import sys

from dataclasses import dataclass

from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QMessageBox, QLabel, QListWidget, QPushButton, QVBoxLayout, QHBoxLayout, QFileDialog, QTextEdit, QLineEdit, QGroupBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize
import pandas as pd
import pickle
import csv
import re

@dataclass
class Settings:
    commentCode: str = None
    statementCode: str = None
    courseNumber: str = None
    roster: pd.DataFrame = None
    header_length: int = 9
    asst_name_idx: int = 4
    asst_points_idx: int = 7
    
    output_directory: str = None

    def save(self, filepath='settings/settings.pkl'):
        with open(filepath, 'wb') as f:
            pickle.dump(self, f)

    @classmethod
    def load(cls, filepath='settings/settings.pkl'):
        try:
            with open(filepath, 'rb') as f:
                return pickle.load(f)
        except FileNotFoundError:
            return cls()

@dataclass
class Data:
    raw_data: list = None
    header_data: list = None
    final_data: pd.DataFrame = None
    name: str = None
    points: list[float] = None
    n_questions: int = None
    qCodes: list[tuple[str, int]] = None
    comment_idx: int = None
    statement_idx: int = None

class MainWindow(QMainWindow):
    def __init__(self):
        # Ensures all initialization code from the inherited class is executed
        super().__init__()

        # Set application spanning styles
        btnMargin = 25
        btnIconHeight = 32
        btnIconSize = QSize(btnIconHeight, btnIconHeight)
        btnHeight = btnIconHeight + btnMargin

        self.setStyleSheet(f"""
            QGroupBox QPushButton {{
                height:  {btnHeight}              
            }}
            QLabel {{
                height: 100px;
                font-family: Arial;
            }}
            QListWidget, QTextEdit, QLineEdit {{
                border: 1px solid grey;
            }}
            QGroupBox {{
                min-width: 200px;
                min-height: 400px;
            }}
        """)

        # Global Variables
        self.settings = Settings().load()

        # Setup output directory
        self.output_dir = 'output'
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        # Initialize main widget
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.setWindowTitle("Cengage Data Scraper")

        # Create the layout
        main_layout = QHBoxLayout()

        # Settings Group
        grpOptions = QGroupBox("Settings")
        optionLayout = QVBoxLayout()
        grpOptions.setLayout(optionLayout)

        lblCourseNumber = QLabel("Course Number")
        self.lnEdtCourseNumber = QLineEdit(f'{self.settings.courseNumber if self.settings.courseNumber else ""}')
        lblStatementCode = QLabel("Statement Code")
        self.lnEdtStatementCode = QLineEdit(f'{self.settings.statementCode if self.settings.statementCode else ""}')
        lblCommentCode = QLabel("Comment Code")
        self.lnEdtCommentCode = QLineEdit(f'{self.settings.commentCode if self.settings.commentCode else ""}')
        self.btnSaveSettings = QPushButton("Save Settings", enabled=True, clicked=self.save_settings)
        
        optionLayout.addWidget(lblCourseNumber)
        optionLayout.addWidget(self.lnEdtCourseNumber)
        optionLayout.addWidget(lblStatementCode)
        optionLayout.addWidget(self.lnEdtStatementCode)
        optionLayout.addWidget(lblCommentCode)
        optionLayout.addWidget(self.lnEdtCommentCode)
        optionLayout.addStretch()
        optionLayout.addWidget(self.btnSaveSettings)

        # Roster Group
        grpRoster = QGroupBox("Current Roster")
        rosterLayout = QVBoxLayout()
        grpRoster.setLayout(rosterLayout)

        self.listRoster = QListWidget()
        self.btnRoster = QPushButton("Load Roster", enabled=True)

        rosterLayout.addWidget(self.listRoster)
        rosterLayout.addWidget(self.btnRoster)
        
        # Data Group
        grpData = QGroupBox("Data")
        dataLayout = QVBoxLayout()
        grpData.setLayout(dataLayout)

        self.txtData = QTextEdit()
        self.txtData.setMinimumWidth(500)
        self.txtData.setLineWrapMode(QTextEdit.NoWrap)
        self.btnLoadData = QPushButton("Load Data", enabled=True, clicked=self.load_data_file)
        self.btnExportData = QPushButton("Export Data", enabled=False)
        self.btnLoadData.setIcon(QIcon('resources/icons/parse.ico'))
        self.btnLoadData.setIconSize(btnIconSize)
        self.btnExportData.setIcon(QIcon('resources/icons/export.png'))
        self.btnExportData.setIconSize(btnIconSize)

        dataLayout.addWidget(self.txtData)
        tempHLayout = QHBoxLayout()
        tempHLayout.addWidget(self.btnLoadData)
        tempHLayout.addWidget(self.btnExportData)
        dataLayout.addLayout(tempHLayout)

        # Connect buttons to functions
        self.btnRoster.clicked.connect(self.setup_roster)

        # Add widgets to the grid
        main_layout.addWidget(grpOptions, stretch=1)
        main_layout.addWidget(grpRoster, stretch=1)
        main_layout.addWidget(grpData, stretch=3)

        # Set the layout for the GUI
        main_widget.setLayout(main_layout)

        self.populate_roster()

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

    def populate_roster(self):
        if self.settings.roster is None: return
        for name, section in zip(self.settings.roster['Cadet Name'], self.settings.roster['Section']):
            self.listRoster.addItem(f'{name} ({section})')
    
    def populate_data_view(self, asst_data: Data):
        self.txtData.setHtml(asst_data.final_data.to_html())

    def process_names(self, text):
        match = re.search(r',[^ ]+', text)
        text = text[:match.end()] if match else text

        return text

    def save_settings(self):
        self.settings.courseNumber = self.lnEdtCourseNumber.text() if self.lnEdtCourseNumber.text() != '' else None
        self.settings.commentCode = self.lnEdtCommentCode.text() if self.lnEdtCommentCode.text() != '' else None
        self.settings.statementCode = self.lnEdtStatementCode.text() if self.lnEdtStatementCode.text() != '' else None
        
        self.settings.save()

    def setup_roster(self):
        # Verify Data
        if self.settings.courseNumber is None: 
            self.show_message('You must enter course information before continuing')
            return

        # Verify intent
        if self.settings.roster is not None and not self.confirm_action('This will reset the current roster, do you wish to continue?'): return

        # Open file dialog
        file_path = self.open_file_dialog("Excel Files (*.xlsx *.xls)")
        if not file_path: 
            return

        try:
            df = pd.read_excel(file_path, skiprows=1)

            df["Course Number"] = df["Course Number"].str.strip()
            df = df[df["Course Number"] == f'{self.settings.courseNumber}'][["Section", "Email", "Cadet Name"]]
            df["Cadet Name"] = df["Cadet Name"].str.strip().map(self.process_names)

            self.settings.roster = df
            self.settings.save()
            self.populate_roster()

        except Exception as e:
            self.show_error(f"Failed to load file\n{e}")
    
    def load_data_file(self):
        file_path = self.open_file_dialog("CSV Files (*.csv)")
        if not file_path:
            return
        
        with open(file_path, 'r') as f:
            data = f.readlines()
            data = [x.strip() for x in data]
            header = data[:self.settings.header_length]
            body = data[self.settings.header_length:]
        
        asst_data = self.process_header(Data(raw_data=body, header_data=header))
        asst_data = self.parse_data(asst_data)
        
        self.populate_data_view(asst_data)

    def process_header(self, asst_data: Data):
        header = asst_data.header_data

        points = [float(x) for x in header[self.settings.asst_points_idx].split(',') if x != '' and x != 'Points']
        name = header[self.settings.asst_name_idx].split(',')[1]
        qCodes = [(x, j) for j, x in enumerate(header[6].split(',')) if x.isdecimal()]

        for code, j in qCodes:
            if code == self.settings.commentCode:
                asst_data.comment_idx = j
            if code == self.settings.statementCode:
                asst_data.statement_idx = j

        asst_data.name = name
        asst_data.points = points
        asst_data.n_questions = len(points)
        asst_data.qCodes = qCodes

        return asst_data
    
    def parse_data(self, data: Data):
        csvReader = csv.reader(data.raw_data)
        students: pd.DataFrame = self.settings.roster
        result = []
        email = ''
        name = ''
        comment = None
        statement = None

        for i, row in enumerate(csvReader):
            if i%2 == 1:
                filter = students['Email'].isin([email])
                if filter.any():
                    section = students[filter]['Section'].iloc[0]
                    nRow = [name, email, section]
                    
                    for j, entry in enumerate(row):
                        if j >= 3:
                            nRow.append(float(entry))
                    
                    if data.comment_idx:
                        nRow[data.comment_idx] = comment if comment else ''
                    if data.statement_idx:
                        nRow[data.statement_idx] = statement if statement else ''

                    result.append(nRow)          
            else:
                email = row[1].split('@usafa')[0]
                name = self.process_names(row[0])
                comment = row[data.comment_idx] if data.comment_idx else None
                statement = row[data.statement_idx] if data.statement_idx else None

        header = ['Name','Email','Section'] + [f'Q{x + 1}' for x in range(data.n_questions + 1)]

        if data.statement_idx:
            header[-1] = 'Statement'
        if data.comment_idx:
            header[-2 if data.statement_idx else -1] = 'Comment'

        data.final_data = pd.DataFrame(result, columns=header)

        return data

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())