
import sys
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QApplication

try:
    from Visualization_ import ChartDialog, DashboardDialog, VisualizationManager
    print("✅ Visualization_ module imported successfully")
except ImportError as e:
    print(f"❌ Failed to import Visualization_ module: {e}")
    sys.exit(1)

if not VisualizationManager.is_available():
    print("❌ VisualizationManager reports dependencies missing")
    sys.exit(1)
else:
    print("✅ VisualizationManager reports dependencies available")

app = QApplication(sys.argv)
df = pd.DataFrame({
    'A': np.random.rand(10),
    'B': np.random.rand(10),
    'C': ['X', 'Y', 'X', 'Y', 'Z', 'X', 'Y', 'Z', 'X', 'Y']
})

try:
    chart_dialog = ChartDialog(df)
    print("✅ ChartDialog instantiated successfully")
    dashboard_dialog = DashboardDialog(df)
    print("✅ DashboardDialog instantiated successfully")
except Exception as e:
    print(f"❌ Failed to instantiate dialogs: {e}")
    sys.exit(1)

print("✅ Visualization verification passed!")
sys.exit(0)
