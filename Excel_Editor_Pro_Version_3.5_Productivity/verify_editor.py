
import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication

try:
    # Attempt to import Editor_ module which caused the error
    from Editor_ import ExcelEditor
    print("[OK] Editor_ module imported successfully")
except ImportError as e:
    print(f"[FAIL] Failed to import Editor_ module: {e}")
    sys.exit(1)
except NameError as e:
    print(f"[FAIL] NameError in Editor_ module: {e}")
    sys.exit(1)
except Exception as e:
    print(f"[FAIL] Unexpected error importing Editor_: {e}")
    sys.exit(1)

# Initialize App (required for QWidget subclass)
app = QApplication(sys.argv)

try:
    editor = ExcelEditor()
    print("[OK] ExcelEditor instantiated successfully")
    
    # Test on_file_loaded
    print("Testing on_file_loaded...")
    df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
    # Mock current_file_path as it is used in on_file_loaded
    editor.current_file_path = "test.csv"
    editor.on_file_loaded(df)
    print("[OK] on_file_loaded executed successfully")
    
    # Test update_table
    print("Testing update_table...")
    editor.update_table()
    print("[OK] update_table executed successfully")
    
except NameError as e:
    print(f"[FAIL] NameError during instantiation/execution: {e}")
    sys.exit(1)
except AttributeError as e:
    print(f"[FAIL] AttributeError during instantiation/execution: {e}")
    sys.exit(1)
except Exception as e:
    print(f"[FAIL] Unexpected error during instantiation/execution: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
    
sys.exit(0)
