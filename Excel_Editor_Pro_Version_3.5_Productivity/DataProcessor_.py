import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal

class DataProcessor(QThread):
    """Background thread for processing large Excel files"""
    finished = pyqtSignal(pd.DataFrame)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)
    sheets_found = pyqtSignal(list)

    def __init__(self, file_path, sheet_name=None):
        super().__init__()
        self.file_path = file_path
        self.sheet_name = sheet_name

    def run(self):
        try:
            self.progress.emit("Loading file...")
            if self.file_path.endswith('.csv'):
                df = pd.read_csv(self.file_path, low_memory=False)
                self.progress.emit("File loaded successfully!")
                self.finished.emit(df)
            else:
                excel_file = pd.ExcelFile(self.file_path)
                sheet_names = excel_file.sheet_names
                
                if len(sheet_names) > 1 and self.sheet_name is None:
                    self.sheets_found.emit(sheet_names)
                    return
                
                sheet_to_load = self.sheet_name if self.sheet_name else sheet_names[0]
                df = pd.read_excel(self.file_path, sheet_name=sheet_to_load, engine='openpyxl')
                
                self.progress.emit("File loaded successfully!")
                self.finished.emit(df)
        except Exception as e:
            self.error.emit(f"Error loading file: {str(e)}")
    
    def get_settings(self):
        return {
            'file_type': self.file_type_combo.currentText(),
            'rows': self.rows_spinbox.value(),
            'columns': self.cols_spinbox.value(),
            'include_headers': self.include_headers_cb.isChecked()
        }