from PyQt5.QtCore import QSettings
from PyQt5.QtGui import QColor

def apply_custom_theme(app, bg_color="#313B2F", accent_color="#FBA002"):
    """Apply custom theme with specified colors"""
    try:
        app.setStyle('Fusion')
        
        palette = app.palette()
        
        # Convert hex to RGB
        bg_qcolor = QColor(bg_color)
        accent_qcolor = QColor(accent_color)
        
        # Create darker shade for base
        darker_bg = bg_qcolor.darker(120)
        
        palette.setColor(palette.Window, bg_qcolor)
        palette.setColor(palette.WindowText, QColor(255, 255, 255))
        palette.setColor(palette.Base, darker_bg)
        palette.setColor(palette.AlternateBase, bg_qcolor)
        palette.setColor(palette.ToolTipBase, bg_qcolor)
        palette.setColor(palette.ToolTipText, QColor(255, 255, 255))
        palette.setColor(palette.Text, QColor(255, 255, 255))
        palette.setColor(palette.Button, bg_qcolor)
        palette.setColor(palette.ButtonText, QColor(255, 255, 255))
        palette.setColor(palette.Highlight, accent_qcolor)
        palette.setColor(palette.HighlightedText, QColor(0, 0, 0))
        
        app.setPalette(palette)
        
        # Get lighter shade for button background
        lighter_bg = bg_qcolor.lighter(110).name()
        darker_bg_hex = darker_bg.name()
        
        stylesheet = f"""
        QMainWindow {{
            background-color: {bg_color};
            color: #ffffff;
        }}
        
        QPushButton {{
            background-color: {lighter_bg};
            color: #ffffff;
            border: 1px solid {accent_color};
            padding: 3px 12px;
            border-radius: 3px;
        }}
        
        QPushButton:hover {{
            background-color: {accent_color};
            color: #000000;
            border-color: {accent_color};
        }}
        
        QPushButton:disabled {{
            background-color: {darker_bg_hex};
            color: #7f7f7f;
            border-color: #555555;
        }}
        
        QTableWidget {{
            background-color: {darker_bg_hex};
            color: #ffffff;
            gridline-color: {lighter_bg};
            selection-background-color: {accent_color};
            selection-color: #000000;
        }}
        
        QTableWidget::item:selected {{
            background-color: {accent_color};
            color: #000000;
        }}
        
        QHeaderView::section {{
            background-color: {lighter_bg};
            color: {accent_color};
            border: 1px solid {lighter_bg};
            padding: 4px;
            font-weight: bold;
        }}
        
        QLineEdit, QComboBox, QTextEdit {{
            background-color: {darker_bg_hex};
            color: #ffffff;
            border: 1px solid {lighter_bg};
            padding: 4px;
            border-radius: 3px;
        }}
        
        QLineEdit:focus, QComboBox:focus, QTextEdit:focus {{
            border: 2px solid {accent_color};
        }}
        
        QComboBox::drop-down {{
            border: 1px solid {lighter_bg};
            background-color: {lighter_bg};
        }}
        
        QComboBox::down-arrow {{
            image: none;
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            border-top: 6px solid {accent_color};
            margin-right: 5px;
        }}
        
        QGroupBox {{
            color: {accent_color};
            border: 2px solid {lighter_bg};
            border-radius: 5px;
            margin-top: 10px;
            font-weight: bold;
        }}
        
        QGroupBox::title {{
            color: {accent_color};
            subcontrol-origin: margin;
            padding: 0 5px;
        }}
        
        QStatusBar {{
            background-color: {bg_color};
            color: #ffffff;
        }}
        
        QMenuBar {{
            background-color: {bg_color};
            color: #ffffff;
        }}
        
        QMenuBar::item:selected {{
            background-color: {accent_color};
            color: #000000;
        }}
        
        QMenu {{
            background-color: {bg_color};
            color: #ffffff;
            border: 1px solid {accent_color};
        }}
        
        QMenu::item:selected {{
            background-color: {accent_color};
            color: #000000;
        }}
        
        QToolBar {{
            background-color: {lighter_bg};
            border: 1px solid {lighter_bg};
            spacing: 3px;
        }}
        
        QScrollBar:vertical {{
            background-color: {bg_color};
            width: 12px;
        }}
        
        QScrollBar::handle:vertical {{
            background-color: {accent_color};
            min-height: 20px;
            border-radius: 3px;
        }}
        
        QScrollBar::handle:vertical:hover {{
            background-color: {accent_color};
            opacity: 0.8;
        }}
        
        QScrollBar:horizontal {{
            background-color: {bg_color};
            height: 12px;
        }}
        
        QScrollBar::handle:horizontal {{
            background-color: {accent_color};
            min-width: 20px;
            border-radius: 6px;
        }}
        
        QScrollBar::handle:horizontal:hover {{
            background-color: {accent_color};
            opacity: 0.8;
        }}
        
        QDialog {{
            background-color: {bg_color};
            color: #ffffff;
        }}
        """
        
        app.setStyleSheet(stylesheet)
        
    except Exception as e:
        print(f"Error applying custom theme: {e}")


def apply_dark_theme(app):
    """Legacy function - redirects to custom theme"""
    settings = QSettings("DataTools", "ExcelEditor")
    bg_color = settings.value("theme/bg_color", "#313B2F")
    accent_color = settings.value("theme/accent_color", "#FBA002")
    apply_custom_theme(app, bg_color, accent_color)