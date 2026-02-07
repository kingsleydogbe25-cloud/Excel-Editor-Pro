"""
Quick Actions Context Menu
Provides right-click context menu with common actions
"""

from PyQt5.QtWidgets import QMenu, QAction, QMessageBox, QInputDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence
import pandas as pd


class QuickActionsMenu:
    """Manages right-click context menu for quick actions"""
    
    def __init__(self, parent):
        self.parent = parent
        self.setup_context_menu()
    
    def setup_context_menu(self):
        """Setup context menu for table widget"""
        self.parent.table_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.parent.table_widget.customContextMenuRequested.connect(self.show_context_menu)
    
    def show_context_menu(self, position):
        """Show context menu at cursor position"""
        menu = QMenu()
        
        # Get current selection
        selected_items = self.parent.table_widget.selectedItems()
        current_row = self.parent.table_widget.currentRow()
        current_col = self.parent.table_widget.currentColumn()
        
        if not selected_items:
            return
        
        # Cell Actions
        cell_menu = menu.addMenu("📋 Cell Actions")
        
        copy_action = QAction("Copy", self.parent)
        copy_action.setShortcut(QKeySequence.Copy)
        copy_action.triggered.connect(self.copy_cells)
        cell_menu.addAction(copy_action)
        
        paste_action = QAction("Paste", self.parent)
        paste_action.setShortcut(QKeySequence.Paste)
        paste_action.triggered.connect(self.paste_cells)
        cell_menu.addAction(paste_action)
        
        cut_action = QAction("Cut", self.parent)
        cut_action.setShortcut(QKeySequence.Cut)
        cut_action.triggered.connect(self.cut_cells)
        cell_menu.addAction(cut_action)
        
        cell_menu.addSeparator()
        
        clear_action = QAction("Clear Content", self.parent)
        clear_action.triggered.connect(self.clear_cells)
        cell_menu.addAction(clear_action)
        
        fill_down_action = QAction("Fill Down", self.parent)
        fill_down_action.triggered.connect(self.fill_down)
        cell_menu.addAction(fill_down_action)
        
        # Row Actions
        menu.addSeparator()
        row_menu = menu.addMenu("📊 Row Actions")
        
        insert_row_above = QAction("Insert Row Above", self.parent)
        insert_row_above.triggered.connect(lambda: self.insert_row(current_row))
        row_menu.addAction(insert_row_above)
        
        insert_row_below = QAction("Insert Row Below", self.parent)
        insert_row_below.triggered.connect(lambda: self.insert_row(current_row + 1))
        row_menu.addAction(insert_row_below)
        
        delete_row_action = QAction("Delete Row", self.parent)
        delete_row_action.setShortcut(QKeySequence.Delete)
        delete_row_action.triggered.connect(self.delete_row)
        row_menu.addAction(delete_row_action)
        
        row_menu.addSeparator()
        
        duplicate_row_action = QAction("Duplicate Row", self.parent)
        duplicate_row_action.triggered.connect(self.duplicate_row)
        row_menu.addAction(duplicate_row_action)
        
        # Column Actions
        menu.addSeparator()
        col_menu = menu.addMenu("📑 Column Actions")
        
        insert_col_left = QAction("Insert Column Left", self.parent)
        insert_col_left.triggered.connect(lambda: self.insert_column(current_col))
        col_menu.addAction(insert_col_left)
        
        insert_col_right = QAction("Insert Column Right", self.parent)
        insert_col_right.triggered.connect(lambda: self.insert_column(current_col + 1))
        col_menu.addAction(insert_col_right)
        
        delete_col_action = QAction("Delete Column", self.parent)
        delete_col_action.triggered.connect(self.delete_column)
        col_menu.addAction(delete_col_action)
        
        col_menu.addSeparator()
        
        rename_col_action = QAction("Rename Column", self.parent)
        rename_col_action.triggered.connect(self.rename_column)
        col_menu.addAction(rename_col_action)
        
        # Data Operations
        menu.addSeparator()
        data_menu = menu.addMenu("🔧 Data Operations")
        
        sort_asc_action = QAction("Sort Ascending", self.parent)
        sort_asc_action.triggered.connect(lambda: self.sort_by_column(True))
        data_menu.addAction(sort_asc_action)
        
        sort_desc_action = QAction("Sort Descending", self.parent)
        sort_desc_action.triggered.connect(lambda: self.sort_by_column(False))
        data_menu.addAction(sort_desc_action)
        
        data_menu.addSeparator()
        
        filter_action = QAction("Filter by Value", self.parent)
        filter_action.triggered.connect(self.filter_by_value)
        data_menu.addAction(filter_action)
        
        # Statistics
        menu.addSeparator()
        stats_action = QAction("📈 Column Statistics", self.parent)
        stats_action.triggered.connect(self.show_column_stats)
        menu.addAction(stats_action)
        
        # Show menu
        menu.exec_(self.parent.table_widget.viewport().mapToGlobal(position))
    
    def copy_cells(self):
        """Copy selected cells to clipboard"""
        try:
            selected_items = self.parent.table_widget.selectedItems()
            if not selected_items:
                return
            
            # Get selection range
            rows = sorted(set(item.row() for item in selected_items))
            cols = sorted(set(item.column() for item in selected_items))
            
            # Build text to copy
            text_lines = []
            for row in rows:
                row_data = []
                for col in cols:
                    item = self.parent.table_widget.item(row, col)
                    row_data.append(item.text() if item else "")
                text_lines.append("\t".join(row_data))
            
            # Copy to clipboard
            from PyQt5.QtWidgets import QApplication
            clipboard = QApplication.clipboard()
            clipboard.setText("\n".join(text_lines))
            
            self.parent.statusBar().showMessage("Copied to clipboard", 2000)
            
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to copy: {str(e)}")
    
    def paste_cells(self):
        """Paste from clipboard"""
        try:
            from PyQt5.QtWidgets import QApplication
            clipboard = QApplication.clipboard()
            text = clipboard.text()
            
            if not text:
                return
            
            current_row = self.parent.table_widget.currentRow()
            current_col = self.parent.table_widget.currentColumn()
            
            if current_row < 0 or current_col < 0:
                return
            
            # Save state for undo
            if hasattr(self.parent, 'undo_redo_manager'):
                self.parent.undo_redo_manager.save_state("Paste")
            
            # Parse clipboard data
            lines = text.split("\n")
            for i, line in enumerate(lines):
                if not line.strip():
                    continue
                
                cells = line.split("\t")
                for j, cell_value in enumerate(cells):
                    target_row = current_row + i
                    target_col = current_col + j
                    
                    if (target_row < self.parent.table_widget.rowCount() and
                        target_col < self.parent.table_widget.columnCount()):
                        
                        from PyQt5.QtWidgets import QTableWidgetItem
                        self.parent.table_widget.setItem(
                            target_row, target_col,
                            QTableWidgetItem(cell_value)
                        )
            
            self.parent.statusBar().showMessage("Pasted from clipboard", 2000)
            
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to paste: {str(e)}")
    
    def cut_cells(self):
        """Cut selected cells"""
        self.copy_cells()
        self.clear_cells()
    
    def clear_cells(self):
        """Clear selected cells"""
        try:
            selected_items = self.parent.table_widget.selectedItems()
            if not selected_items:
                return
            
            # Save state for undo
            if hasattr(self.parent, 'undo_redo_manager'):
                self.parent.undo_redo_manager.save_state("Clear Cells")
            
            for item in selected_items:
                item.setText("")
            
            self.parent.statusBar().showMessage("Cells cleared", 2000)
            
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to clear cells: {str(e)}")
    
    def fill_down(self):
        """Fill down from first selected cell"""
        try:
            selected_items = self.parent.table_widget.selectedItems()
            if len(selected_items) < 2:
                return
            
            # Save state for undo
            if hasattr(self.parent, 'undo_redo_manager'):
                self.parent.undo_redo_manager.save_state("Fill Down")
            
            # Get first item value
            first_item = selected_items[0]
            fill_value = first_item.text()
            
            # Fill all selected cells
            for item in selected_items[1:]:
                if item.column() == first_item.column():
                    item.setText(fill_value)
            
            self.parent.statusBar().showMessage("Filled down", 2000)
            
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to fill down: {str(e)}")
    
    def insert_row(self, position):
        """Insert row at position"""
        try:
            if self.parent.df is None:
                return
            
            # Save state for undo
            if hasattr(self.parent, 'undo_redo_manager'):
                self.parent.undo_redo_manager.save_state("Insert Row")
            
            # Create empty row
            new_row = pd.DataFrame([[None] * len(self.parent.df.columns)],
                                  columns=self.parent.df.columns)
            
            # Insert row
            self.parent.df = pd.concat([
                self.parent.df.iloc[:position],
                new_row,
                self.parent.df.iloc[position:]
            ]).reset_index(drop=True)
            
            self.parent.update_table()
            self.parent.is_modified = True
            
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to insert row: {str(e)}")
    
    def delete_row(self):
        """Delete selected row"""
        self.parent.delete_row()
    
    def duplicate_row(self):
        """Duplicate selected row"""
        try:
            current_row = self.parent.table_widget.currentRow()
            if current_row < 0 or self.parent.df is None:
                return
            
            # Save state for undo
            if hasattr(self.parent, 'undo_redo_manager'):
                self.parent.undo_redo_manager.save_state("Duplicate Row")
            
            # Get row data
            row_data = self.parent.df.iloc[current_row:current_row+1].copy()
            
            # Insert duplicate
            self.parent.df = pd.concat([
                self.parent.df.iloc[:current_row+1],
                row_data,
                self.parent.df.iloc[current_row+1:]
            ]).reset_index(drop=True)
            
            self.parent.update_table()
            self.parent.is_modified = True
            
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to duplicate row: {str(e)}")
    
    def insert_column(self, position):
        """Insert column at position"""
        try:
            if self.parent.df is None:
                return
            
            # Get column name
            col_name, ok = QInputDialog.getText(
                self.parent,
                "Insert Column",
                "Enter column name:"
            )
            
            if not ok or not col_name:
                return
            
            # Save state for undo
            if hasattr(self.parent, 'undo_redo_manager'):
                self.parent.undo_redo_manager.save_state("Insert Column")
            
            # Insert column
            self.parent.df.insert(position, col_name, None)
            
            self.parent.update_table()
            self.parent.is_modified = True
            
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to insert column: {str(e)}")
    
    def delete_column(self):
        """Delete selected column"""
        try:
            current_col = self.parent.table_widget.currentColumn()
            if current_col < 0 or self.parent.df is None:
                return
            
            col_name = self.parent.df.columns[current_col]
            
            reply = QMessageBox.question(
                self.parent,
                "Delete Column",
                f"Are you sure you want to delete column '{col_name}'?",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                # Save state for undo
                if hasattr(self.parent, 'undo_redo_manager'):
                    self.parent.undo_redo_manager.save_state("Delete Column")
                
                self.parent.df = self.parent.df.drop(columns=[col_name])
                self.parent.update_table()
                self.parent.is_modified = True
                
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to delete column: {str(e)}")
    
    def rename_column(self):
        """Rename selected column"""
        try:
            current_col = self.parent.table_widget.currentColumn()
            if current_col < 0 or self.parent.df is None:
                return
            
            old_name = self.parent.df.columns[current_col]
            
            new_name, ok = QInputDialog.getText(
                self.parent,
                "Rename Column",
                f"Enter new name for '{old_name}':",
                text=old_name
            )
            
            if ok and new_name and new_name != old_name:
                # Save state for undo
                if hasattr(self.parent, 'undo_redo_manager'):
                    self.parent.undo_redo_manager.save_state("Rename Column")
                
                self.parent.df = self.parent.df.rename(columns={old_name: new_name})
                self.parent.update_table()
                self.parent.is_modified = True
                
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to rename column: {str(e)}")
    
    def sort_by_column(self, ascending=True):
        """Sort by selected column"""
        try:
            current_col = self.parent.table_widget.currentColumn()
            if current_col < 0 or self.parent.df is None:
                return
            
            col_name = self.parent.df.columns[current_col]
            
            # Save state for undo
            if hasattr(self.parent, 'undo_redo_manager'):
                self.parent.undo_redo_manager.save_state("Sort Column")
            
            self.parent.df = self.parent.df.sort_values(by=col_name, ascending=ascending)
            self.parent.update_table()
            self.parent.is_modified = True
            
            direction = "ascending" if ascending else "descending"
            self.parent.statusBar().showMessage(f"Sorted by {col_name} ({direction})", 2000)
            
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to sort: {str(e)}")
    
    def filter_by_value(self):
        """Filter by current cell value"""
        try:
            current_item = self.parent.table_widget.currentItem()
            if not current_item or self.parent.df is None:
                return
            
            current_col = self.parent.table_widget.currentColumn()
            col_name = self.parent.df.columns[current_col]
            value = current_item.text()
            
            # Set filter
            self.parent.filter_column_combo.setCurrentText(col_name)
            self.parent.filter_input.setText(value)
            self.parent.apply_filter()
            
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to filter: {str(e)}")
    
    def show_column_stats(self):
        """Show statistics for selected column"""
        try:
            current_col = self.parent.table_widget.currentColumn()
            if current_col < 0 or self.parent.df is None:
                return
            
            col_name = self.parent.df.columns[current_col]
            col_data = self.parent.df[col_name]
            
            # Calculate statistics
            stats = []
            stats.append(f"Column: {col_name}")
            stats.append(f"Total Values: {len(col_data)}")
            stats.append(f"Non-Null: {col_data.notna().sum()}")
            stats.append(f"Null: {col_data.isna().sum()}")
            stats.append(f"Unique: {col_data.nunique()}")
            
            # Numeric statistics
            if pd.api.types.is_numeric_dtype(col_data):
                stats.append(f"\nNumeric Statistics:")
                stats.append(f"Mean: {col_data.mean():.2f}")
                stats.append(f"Median: {col_data.median():.2f}")
                stats.append(f"Std Dev: {col_data.std():.2f}")
                stats.append(f"Min: {col_data.min():.2f}")
                stats.append(f"Max: {col_data.max():.2f}")
            
            QMessageBox.information(
                self.parent,
                "Column Statistics",
                "\n".join(stats)
            )
            
        except Exception as e:
            QMessageBox.critical(self.parent, "Error", f"Failed to show statistics: {str(e)}")
