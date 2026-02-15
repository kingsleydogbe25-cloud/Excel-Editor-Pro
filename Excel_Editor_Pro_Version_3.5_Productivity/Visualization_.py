from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLabel, QComboBox, QWidget, QMessageBox, QGroupBox,
                             QScrollArea, QTabWidget, QListWidget, QGridLayout)
from PyQt5.QtCore import Qt
try:
    import matplotlib
    matplotlib.use('Qt5Agg')
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.figure import Figure
    import matplotlib.pyplot as plt
    import seaborn as sns
    import pandas as pd
    import numpy as np
    VISUALIZATION_AVAILABLE = True
except ImportError:
    VISUALIZATION_AVAILABLE = False

class VisualizationManager:
    """Helper to check availability and launch dialogs"""
    @staticmethod
    def is_available():
        return VISUALIZATION_AVAILABLE

    @staticmethod
    def get_missing_deps_message():
        return "Visualization features require 'matplotlib' and 'seaborn'.\nPlease install: pip install matplotlib seaborn"

class ChartDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.df = df
        self.setWindowTitle("📊 Create Chart")
        self.resize(1000, 700)
        
        # Filter for numeric and categorical columns
        self.numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        self.all_cols = df.columns.tolist()
        
        self.init_ui()
        
    def init_ui(self):
        main_layout = QHBoxLayout()
        
        # Left Panel: Controls
        control_panel = QGroupBox("Chart Settings")
        control_panel.setMaximumWidth(300)
        control_layout = QVBoxLayout()
        
        # Chart Type Selection
        control_layout.addWidget(QLabel("Chart Type:"))
        self.chart_type_combo = QComboBox()
        self.chart_type_combo.addItems(["Histogram", "Scatter Plot", "Box Plot", "Pie Chart", "Bar Chart", "Line Plot"])
        self.chart_type_combo.currentTextChanged.connect(self.update_controls)
        control_layout.addWidget(self.chart_type_combo)
        
        # X Axis Selection
        control_layout.addWidget(QLabel("X Axis / Category:"))
        self.x_col_combo = QComboBox()
        self.x_col_combo.addItems(self.all_cols)
        control_layout.addWidget(self.x_col_combo)
        
        # Y Axis Selection
        self.y_label = QLabel("Y Axis / Value:")
        control_layout.addWidget(self.y_label)
        self.y_col_combo = QComboBox()
        self.y_col_combo.addItems(self.numeric_cols)
        control_layout.addWidget(self.y_col_combo)
        
        # Color/Hue Selection (Optional)
        control_layout.addWidget(QLabel("Color Grouping (Hue):"))
        self.hue_col_combo = QComboBox()
        self.hue_col_combo.addItem("None")
        self.hue_col_combo.addItems(self.all_cols)
        control_layout.addWidget(self.hue_col_combo)
        
        # Plot Button
        self.plot_btn = QPushButton("📉 Generate Chart")
        self.plot_btn.clicked.connect(self.plot_chart)
        self.plot_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8px;")
        control_layout.addWidget(self.plot_btn)
        
        control_layout.addStretch()
        control_panel.setLayout(control_layout)
        main_layout.addWidget(control_panel)
        
        # Right Panel: Plot Area
        self.plot_area = QGroupBox("Chart Preview")
        plot_layout = QVBoxLayout()
        
        # Matplotlib Figure
        self.figure = Figure(figsize=(8, 6), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        plot_layout.addWidget(self.canvas)
        
        self.plot_area.setLayout(plot_layout)
        main_layout.addWidget(self.plot_area)
        
        self.setLayout(main_layout)
        
        # Initial control state update
        self.update_controls(self.chart_type_combo.currentText())
        
    def update_controls(self, chart_type):
        """Enable/Disable controls based on chart type"""
        if chart_type == "Histogram":
            self.y_col_combo.setEnabled(False)
            self.hue_col_combo.setEnabled(True)
            self.y_label.setText("Y Axis (Frequency - Auto)")
        elif chart_type == "Pie Chart":
            self.y_col_combo.setEnabled(False) # Count of categories
            self.hue_col_combo.setEnabled(False)
            self.y_label.setText("Value (Count - Auto)")
        elif chart_type == "Box Plot":
            self.y_col_combo.setEnabled(True)
            self.hue_col_combo.setEnabled(True)
        else: # Scatter, Bar, Line
            self.y_col_combo.setEnabled(True)
            self.hue_col_combo.setEnabled(True)
            self.y_label.setText("Y Axis / Value:")

    def plot_chart(self):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        chart_type = self.chart_type_combo.currentText()
        x_col = self.x_col_combo.currentText()
        y_col = self.y_col_combo.currentText()
        hue_col = self.hue_col_combo.currentText()
        hue = hue_col if hue_col != "None" else None
        
        try:
            if chart_type == "Histogram":
                sns.histplot(data=self.df, x=x_col, hue=hue, kde=True, ax=ax)
                ax.set_title(f"Histogram of {x_col}")
                
            elif chart_type == "Scatter Plot":
                if x_col not in self.numeric_cols or y_col not in self.numeric_cols:
                    if x_col not in self.numeric_cols: # Allow categorical X for swarm like plots? For now stick to strict scatter
                        QMessageBox.warning(self, "Warning", "For Scatter Plot, both X and Y must be numeric (or convertable).")
                        # Proceeding anyway as matplotlib might handle it, but seaborn style scatter usually implies numeric
                sns.scatterplot(data=self.df, x=x_col, y=y_col, hue=hue, ax=ax)
                ax.set_title(f"{y_col} vs {x_col}")
                
            elif chart_type == "Box Plot":
                sns.boxplot(data=self.df, x=x_col, y=y_col if self.y_col_combo.isEnabled() else None, hue=hue, ax=ax)
                ax.set_title(f"Box Plot of {x_col}")
                
            elif chart_type == "Pie Chart":
                # Aggregate data for pie chart
                data = self.df[x_col].value_counts()
                ax.pie(data, labels=data.index, autopct='%1.1f%%', startangle=90)
                ax.set_title(f"Distribution of {x_col}")
                
            elif chart_type == "Bar Chart":
                if self.y_col_combo.isEnabled():
                    # If Y is specified, plot X vs Y (e.g. Categories vs Values)
                    # We usually want to aggregate if there are duplicates for X
                    sns.barplot(data=self.df, x=x_col, y=y_col, hue=hue, ax=ax)
                else:
                    sns.countplot(data=self.df, x=x_col, hue=hue, ax=ax)
                ax.set_title(f"Bar Chart of {x_col}")
                
            elif chart_type == "Line Plot":
                 sns.lineplot(data=self.df, x=x_col, y=y_col, hue=hue, ax=ax)
                 ax.set_title(f"Line Plot: {y_col} over {x_col}")
            
            plt.tight_layout()
            self.canvas.draw()
            
        except Exception as e:
            QMessageBox.critical(self, "Plot Error", f"Could not generate plot: {str(e)}")


class DashboardDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.df = df
        self.setWindowTitle("🚀 Data Dashboard")
        self.resize(1200, 800)
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Info Header
        info_layout = QHBoxLayout()
        info_layout.addWidget(QLabel(f"<h2>Dashboard: {len(self.df)} Rows, {len(self.df.columns)} Columns</h2>"))
        layout.addLayout(info_layout)
        
        # Scroll area for dashboard content
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content_widget = QWidget()
        self.grid = QGridLayout()
        content_widget.setLayout(self.grid)
        scroll.setWidget(content_widget)
        layout.addWidget(scroll)
        
        self.generate_dashboard_widgets()
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)
        
    def generate_dashboard_widgets(self):
        """Generate widgets for the dashboard grid"""
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns.tolist()
        categorical_cols = self.df.select_dtypes(include=['object', 'category']).columns.tolist()
        
        row, col = 0, 0
        max_cols = 2
        
        # 1. Numeric Distributions (Histograms) - Top 4
        for col_name in numeric_cols[:4]:
            fig = Figure(figsize=(5, 4), dpi=100)
            canvas = FigureCanvas(fig)
            ax = fig.add_subplot(111)
            try:
                sns.histplot(data=self.df, x=col_name, kde=True, ax=ax, color='skyblue')
                ax.set_title(f"Distribution: {col_name}")
                plt.tight_layout()
                self.grid.addWidget(canvas, row, col)
                col += 1
                if col >= max_cols:
                    col = 0
                    row += 1
            except: pass

        # 2. Categorical Counts (Count Plots) - Top 2
        for col_name in categorical_cols[:2]:
            fig = Figure(figsize=(5, 4), dpi=100)
            canvas = FigureCanvas(fig)
            ax = fig.add_subplot(111)
            try:
                # Limit to top 10 categories to avoid clutter
                top_cats = self.df[col_name].value_counts().nlargest(10).index
                sns.countplot(data=self.df[self.df[col_name].isin(top_cats)], x=col_name, ax=ax, palette='viridis')
                ax.set_title(f"Counts: {col_name}")
                ax.tick_params(axis='x', rotation=45)
                plt.tight_layout()
                self.grid.addWidget(canvas, row, col)
                col += 1
                if col >= max_cols:
                    col = 0
                    row += 1
            except: pass
            
        # 3. Correlation Heatmap (if enough numeric cols)
        if len(numeric_cols) > 1:
            fig = Figure(figsize=(5, 4), dpi=100)
            canvas = FigureCanvas(fig)
            ax = fig.add_subplot(111)
            try:
                corr = self.df[numeric_cols].corr()
                sns.heatmap(corr, annot=True, cmap='coolwarm', fmt=".2f", ax=ax)
                ax.set_title("Correlation Matrix")
                plt.tight_layout()
                self.grid.addWidget(canvas, row, col)
                col += 1
                if col >= max_cols:
                    col = 0
                    row += 1
            except: pass
