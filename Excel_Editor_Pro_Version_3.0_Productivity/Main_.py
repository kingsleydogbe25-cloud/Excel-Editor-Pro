import sys
import traceback
from PyQt5.QtWidgets import (QApplication)
from PyQt5.QtCore import Qt
from SplashScreen_ import SplashScreen
from Editor_ import ExcelEditor
from Theme_ import apply_dark_theme


def main():
    print("Starting Excel Editor...")
    
    try:
        app = QApplication(sys.argv)
        print("QApplication created")
        
        # Create and show splash screen FIRST
        splash = SplashScreen()
        splash.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        splash.show()
        
        # Force display the splash screen
        app.processEvents()
        
        splash.update_progress(0, "Starting application...")
        app.processEvents()
        
        # Simulate loading steps with progress updates
        import time
        
        splash.update_progress(10, "Loading libraries...")
        app.processEvents()
        time.sleep(2.4)
        
        splash.update_progress(25, "Initializing components...")
        app.processEvents()
        time.sleep(1)
        
        splash.update_progress(40, "Loading theme settings...")
        app.processEvents()
        apply_dark_theme(app)
        print("Dark theme applied")
        time.sleep(0.5)
        
        splash.update_progress(55, "Setting up application...")
        app.processEvents()
        app.setApplicationName("Excel Editor Pro")
        app.setApplicationVersion("3.0 - Optimized")
        app.setOrganizationName("Data Tools")
        time.sleep(1.2)
        
        splash.update_progress(70, "Creating main window...")
        app.processEvents()
        print("Creating main window...")
        window = ExcelEditor()
        print("Main window created")
        time.sleep(2.4)
        
        splash.update_progress(85, "Finalizing setup...")
        app.processEvents()
        time.sleep(0.8)
        
        splash.update_progress(100, "Ready!")
        app.processEvents()
        time.sleep(0.5)
        
        # Show main window and close splash
        window.show()
        splash.finish(window)
        print("Window shown")
        
        print("Starting event loop...")
        sys.exit(app.exec_())
        
    except Exception as e:
        print(f"Fatal error in main: {e}")
        traceback.print_exc()
        input("Press Enter to exit...")
        sys.exit(1)


if __name__ == '__main__':
    main()