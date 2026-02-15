from PyQt5.QtWidgets import (QVBoxLayout, QDialog, QDialogButtonBox,  QTextEdit,)

class StatisticsDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Data Statistics")
        self.setModal(True)
        self.resize(600, 500)
        
        layout = QVBoxLayout()
        
        stats_text = QTextEdit()
        stats_text.setReadOnly(True)
        
        try:
            stats_info = []
            stats_info.append(f"Dataset Shape: {df.shape[0]} rows × {df.shape[1]} columns\n")
            stats_info.append("=" * 50 + "\n")
            
            stats_info.append("COLUMN INFORMATION:\n")
            for col in df.columns:
                dtype = str(df[col].dtype)
                null_count = df[col].isnull().sum()
                unique_count = df[col].nunique()
                stats_info.append(f"{col}:\n")
                stats_info.append(f"  - Data Type: {dtype}\n")
                stats_info.append(f"  - Null Values: {null_count}\n")
                stats_info.append(f"  - Unique Values: {unique_count}\n\n")
            
            numerical_cols = df.select_dtypes(include=['number']).columns
            if len(numerical_cols) > 0:
                stats_info.append("NUMERICAL STATISTICS:\n")
                desc_stats = df[numerical_cols].describe()
                stats_info.append(desc_stats.to_string())
                stats_info.append("\n\n")
            
            stats_text.setPlainText("".join(stats_info))
        except Exception as e:
            stats_text.setPlainText(f"Error generating statistics: {str(e)}")
        
        layout.addWidget(stats_text)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)