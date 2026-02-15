

def open_visualization_menu(self):
    """Open a context menu for visualization options"""
    menu = QMenu(self)
    
    chart_action = QAction("📊 Create Chart", self)
    chart_action.triggered.connect(self.create_chart)
    menu.addAction(chart_action)
    
    dashboard_action = QAction("🚀 Dashboard View", self)
    dashboard_action.triggered.connect(self.open_dashboard)
    menu.addAction(dashboard_action)
    
    # Show menu under the button
    menu.exec_(self.visualize_btn.mapToGlobal(self.visualize_btn.rect().bottomLeft()))

def create_chart(self):
    """Open chart creation dialog"""
    if not VisualizationManager.is_available():
        QMessageBox.warning(self, "Missing Dependencies", VisualizationManager.get_missing_deps_message())
        return
        
    if self.df is None or self.df.empty:
        QMessageBox.warning(self, "No Data", "Please load a file with data first.")
        return
        
    try:
        dialog = ChartDialog(self.df, self)
        dialog.exec_()
    except Exception as e:
        QMessageBox.critical(self, "Error", f"Error opening chart dialog: {str(e)}")

def open_dashboard(self):
    """Open dashboard view"""
    if not VisualizationManager.is_available():
        QMessageBox.warning(self, "Missing Dependencies", VisualizationManager.get_missing_deps_message())
        return
        
    if self.df is None or self.df.empty:
        QMessageBox.warning(self, "No Data", "Please load a file with data first.")
        return
        
    try:
        dialog = DashboardDialog(self.df, self)
        dialog.exec_()
    except Exception as e:
        QMessageBox.critical(self, "Error", f"Error opening dashboard: {str(e)}")

