from PyQt5.QtWidgets import(QSplashScreen, QApplication)
from PyQt5.QtGui import QFont, QColor, QPixmap, QPainter, QLinearGradient, QPen, QRadialGradient
from PyQt5.QtCore import Qt, QSettings, QPointF, QRectF
 
class SplashScreen(QSplashScreen):
    """Modern splash screen with loading progress"""
    def __init__(self):
        # Create a pixmap for the splash screen
        pixmap = QPixmap(600, 400)
        pixmap.fill(Qt.transparent)
        
        super().__init__(pixmap)
        self.setWindowFlag(Qt.WindowStaysOnTopHint)
        self.setWindowFlag(Qt.FramelessWindowHint)  # Modern frameless window
        
        # Get saved theme colors or use defaults
        settings = QSettings("DataTools", "ExcelEditor")
        self.bg_color = QColor(settings.value("theme/bg_color", "#313B2F"))
        self.accent_color = QColor(settings.value("theme/accent_color", "#FBA002"))
        
        self.progress = 0
        self.message = "Initializing..."
        
        self.draw_splash()
    
    def draw_splash(self):
        """Draw the modern splash screen with gradient and branding"""
        pixmap = QPixmap(600, 400)
        pixmap.fill(Qt.transparent)
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)  # Smoother rendering
        
        # Create modern gradient background with 3 color stops
        gradient = QLinearGradient(0, 0, 0, 400)
        gradient.setColorAt(0, self.bg_color.lighter(130))
        gradient.setColorAt(0.5, self.bg_color)
        gradient.setColorAt(1, self.bg_color.darker(200))
        
        # Draw rounded rectangle background with larger radius
        painter.setBrush(gradient)
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(0, 0, 600, 400, 15, 15)
        
        # Add subtle decorative elements (modern touch)
        painter.setOpacity(0.05)
        radial = QRadialGradient(QPointF(500, 100), 150)
        radial.setColorAt(0, self.accent_color)
        radial.setColorAt(1, Qt.transparent)
        painter.setBrush(radial)
        painter.drawEllipse(QPointF(500, 100), 150, 150)
        painter.setOpacity(1.0)
        
        # Modern border with glow effect
        painter.setOpacity(0.6)
        border_pen = QPen(self.accent_color, 2)
        painter.setPen(border_pen)
        painter.setBrush(Qt.NoBrush)
        painter.drawRoundedRect(3, 3, 594, 394, 15, 15)
        painter.setOpacity(1.0)
        
        # Draw application title with modern font
        painter.setPen(self.accent_color)
        title_font = QFont("Segoe UI", 46, QFont.Bold)  # Modern font, slightly larger
        title_font.setLetterSpacing(QFont.AbsoluteSpacing, 1)
        painter.setFont(title_font)
        
        # Add text shadow for depth
        painter.setOpacity(0.3)
        painter.drawText(51, 51, 500, 70, Qt.AlignCenter, "Excel Editor Pro")
        painter.setOpacity(1.0)
        painter.drawText(50, 50, 500, 70, Qt.AlignCenter, "Excel Editor Pro")
        
        # Draw subtitle with better contrast
        painter.setPen(QColor(220, 220, 220))
        subtitle_font = QFont("Segoe UI", 13, QFont.Normal)
        painter.setFont(subtitle_font)
        painter.drawText(50, 140, 500, 30, Qt.AlignCenter, "Advanced Data Management Tool")
        
        # Modern version badge instead of plain text
        badge_width = 100
        badge_height = 24
        badge_x = (600 - badge_width) / 2
        badge_y = 175
        
        badge_gradient = QLinearGradient(badge_x, badge_y, badge_x, badge_y + badge_height)
        badge_gradient.setColorAt(0, self.accent_color.darker(110))
        badge_gradient.setColorAt(1, self.accent_color)
        
        painter.setBrush(badge_gradient)
        painter.setPen(QPen(self.accent_color.lighter(120), 1))
        painter.drawRoundedRect(QRectF(badge_x, badge_y, badge_width, badge_height), 12, 12)
        
        painter.setPen(QColor(255, 255, 255))
        version_font = QFont("Segoe UI", 9, QFont.DemiBold)
        painter.setFont(version_font)
        painter.drawText(int(badge_x), int(badge_y), badge_width, badge_height, 
                        Qt.AlignCenter, "VERSION 3.5")
        
        # Modern decorative accent line (thinner, more elegant)
        painter.setOpacity(0.7)
        line_pen = QPen(self.accent_color, 1)
        painter.setPen(line_pen)
        painter.drawLine(120, 220, 480, 220)
        painter.setOpacity(1.0)
        
        # Draw loading message with modern styling
        message_font = QFont("Segoe UI", 11)
        painter.setFont(message_font)
        painter.setPen(QColor(200, 200, 200))
        painter.drawText(50, 245, 500, 30, Qt.AlignCenter, self.message)
        
        # Modern progress bar with rounded ends
        progress_y = 290
        progress_x = 100
        progress_width = 400
        progress_height = 6  # Thinner, more modern
        
        # Progress bar background
        painter.setBrush(QColor(255, 255, 255, 30))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(progress_x, progress_y, progress_width, progress_height, 3, 3)
        
        # Progress bar fill with gradient
        if self.progress > 0:
            fill_width = int(progress_width * (self.progress / 100))
            
            # Gradient fill for modern look
            fill_gradient = QLinearGradient(progress_x, progress_y, 
                                          progress_x + fill_width, progress_y)
            fill_gradient.setColorAt(0, self.accent_color.lighter(110))
            fill_gradient.setColorAt(1, self.accent_color)
            
            painter.setBrush(fill_gradient)
            painter.drawRoundedRect(progress_x, progress_y, fill_width, progress_height, 3, 3)
            
            # Add subtle glow effect on progress
            painter.setOpacity(0.4)
            painter.drawRoundedRect(progress_x, progress_y - 1, fill_width, progress_height + 2, 3, 3)
            painter.setOpacity(1.0)
        
        # Draw progress percentage (modern positioning)
        painter.setPen(self.accent_color.lighter(130))
        percent_font = QFont("Segoe UI", 10, QFont.DemiBold)
        painter.setFont(percent_font)
        painter.drawText(100, 305, 400, 25, Qt.AlignRight | Qt.AlignVCenter, 
                        f"{int(self.progress)}%")
        
        # Draw footer with modern styling
        footer_font = QFont("Segoe UI", 8)
        painter.setFont(footer_font)
        painter.setPen(QColor(160, 160, 160))
        painter.drawText(50, 350, 500, 20, Qt.AlignCenter, "© 2026 SyncNet Technology")
        
        painter.setPen(QColor(130, 130, 130))
        painter.setFont(QFont("Segoe UI", 7))
        painter.drawText(50, 370, 500, 20, Qt.AlignCenter, 
                        "Data Transformation & Advanced Analytics")
        
        painter.end()
        
        self.setPixmap(pixmap)
    
    def update_progress(self, value, message=""):
        """Update the progress bar and message"""
        self.progress = value
        if message:
            self.message = message
        self.draw_splash()
        QApplication.processEvents()
