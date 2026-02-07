"""
Enhanced Undo/Redo System with History Viewer
Provides comprehensive undo/redo functionality with visual history
"""

import copy
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                             QListWidget, QLabel, QMessageBox, QListWidgetItem)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class UndoRedoManager:
    """Manages undo/redo operations with history tracking"""
    
    def __init__(self, parent):
        self.parent = parent
        self.undo_stack = []
        self.redo_stack = []
        self.max_stack_size = 50  # Maximum number of undo states to keep
        self.current_position = -1
    
    def save_state(self, description="Edit"):
        """Save current state to undo stack"""
        try:
            if self.parent.df is None:
                return
            
            # Create a deep copy of the dataframe
            state = {
                'df': self.parent.df.copy(),
                'description': description,
                'file_path': self.parent.current_file_path,
                'format_settings': copy.deepcopy(self.parent.format_settings)
            }
            
            # Clear redo stack when new action is performed
            self.redo_stack.clear()
            
            # Add to undo stack
            self.undo_stack.append(state)
            
            # Limit stack size
            if len(self.undo_stack) > self.max_stack_size:
                self.undo_stack.pop(0)
            
            self.current_position = len(self.undo_stack) - 1
            
        except Exception as e:
            print(f"Error saving undo state: {e}")
    
    def undo(self):
        """Undo last action"""
        try:
            if not self.can_undo():
                self.parent.statusBar().showMessage("Nothing to undo", 2000)
                return False
            
            # Save current state to redo stack
            current_state = {
                'df': self.parent.df.copy(),
                'description': "Current State",
                'file_path': self.parent.current_file_path,
                'format_settings': copy.deepcopy(self.parent.format_settings)
            }
            self.redo_stack.append(current_state)
            
            # Get previous state
            previous_state = self.undo_stack.pop()
            
            # Restore state
            self.parent.df = previous_state['df'].copy()
            self.parent.current_file_path = previous_state['file_path']
            self.parent.format_settings = copy.deepcopy(previous_state['format_settings'])
            
            # Update UI
            self.parent.update_table()
            self.parent.is_modified = True
            
            self.current_position = len(self.undo_stack) - 1
            
            self.parent.statusBar().showMessage(
                f"Undone: {previous_state['description']}", 2000
            )
            
            return True
            
        except Exception as e:
            print(f"Error during undo: {e}")
            QMessageBox.critical(
                self.parent,
                "Undo Error",
                f"Failed to undo:\n{str(e)}"
            )
            return False
    
    def redo(self):
        """Redo last undone action"""
        try:
            if not self.can_redo():
                self.parent.statusBar().showMessage("Nothing to redo", 2000)
                return False
            
            # Get next state
            next_state = self.redo_stack.pop()
            
            # Save current to undo stack
            current_state = {
                'df': self.parent.df.copy(),
                'description': "Undo",
                'file_path': self.parent.current_file_path,
                'format_settings': copy.deepcopy(self.parent.format_settings)
            }
            self.undo_stack.append(current_state)
            
            # Restore state
            self.parent.df = next_state['df'].copy()
            self.parent.current_file_path = next_state['file_path']
            self.parent.format_settings = copy.deepcopy(next_state['format_settings'])
            
            # Update UI
            self.parent.update_table()
            self.parent.is_modified = True
            
            self.current_position = len(self.undo_stack) - 1
            
            self.parent.statusBar().showMessage(
                f"Redone: {next_state['description']}", 2000
            )
            
            return True
            
        except Exception as e:
            print(f"Error during redo: {e}")
            QMessageBox.critical(
                self.parent,
                "Redo Error",
                f"Failed to redo:\n{str(e)}"
            )
            return False
    
    def can_undo(self):
        """Check if undo is available"""
        return len(self.undo_stack) > 0
    
    def can_redo(self):
        """Check if redo is available"""
        return len(self.redo_stack) > 0
    
    def clear(self):
        """Clear all undo/redo history"""
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.current_position = -1
    
    def get_undo_count(self):
        """Get number of available undo actions"""
        return len(self.undo_stack)
    
    def get_redo_count(self):
        """Get number of available redo actions"""
        return len(self.redo_stack)


class UndoRedoHistoryDialog(QDialog):
    """Dialog to view and navigate undo/redo history"""
    
    def __init__(self, parent, undo_redo_manager):
        super().__init__(parent)
        self.parent = parent
        self.undo_redo_manager = undo_redo_manager
        self.init_ui()
        self.load_history()
    
    def init_ui(self):
        self.setWindowTitle("Edit History")
        self.setGeometry(200, 200, 500, 600)
        
        layout = QVBoxLayout()
        
        # Info section
        info_layout = QHBoxLayout()
        
        self.undo_count_label = QLabel()
        info_layout.addWidget(self.undo_count_label)
        
        info_layout.addStretch()
        
        self.redo_count_label = QLabel()
        info_layout.addWidget(self.redo_count_label)
        
        layout.addLayout(info_layout)
        
        # History lists
        lists_layout = QHBoxLayout()
        
        # Undo history
        undo_layout = QVBoxLayout()
        undo_layout.addWidget(QLabel("Undo History (Most Recent First):"))
        
        self.undo_list = QListWidget()
        self.undo_list.itemDoubleClicked.connect(self.undo_to_item)
        undo_layout.addWidget(self.undo_list)
        
        lists_layout.addLayout(undo_layout)
        
        # Redo history
        redo_layout = QVBoxLayout()
        redo_layout.addWidget(QLabel("Redo History:"))
        
        self.redo_list = QListWidget()
        self.redo_list.itemDoubleClicked.connect(self.redo_to_item)
        redo_layout.addWidget(self.redo_list)
        
        lists_layout.addLayout(redo_layout)
        
        layout.addLayout(lists_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        undo_btn = QPushButton("Undo")
        undo_btn.clicked.connect(self.perform_undo)
        button_layout.addWidget(undo_btn)
        
        redo_btn = QPushButton("Redo")
        redo_btn.clicked.connect(self.perform_redo)
        button_layout.addWidget(redo_btn)
        
        clear_btn = QPushButton("Clear History")
        clear_btn.clicked.connect(self.clear_history)
        button_layout.addWidget(clear_btn)
        
        button_layout.addStretch()
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def load_history(self):
        """Load undo/redo history into lists"""
        self.undo_list.clear()
        self.redo_list.clear()
        
        # Load undo history (reverse order - most recent first)
        for i, state in enumerate(reversed(self.undo_redo_manager.undo_stack)):
            item = QListWidgetItem(f"{i+1}. {state['description']}")
            
            # Make current state bold
            if i == 0:
                font = item.font()
                font.setBold(True)
                item.setFont(font)
            
            self.undo_list.addItem(item)
        
        # Load redo history
        for i, state in enumerate(reversed(self.undo_redo_manager.redo_stack)):
            item = QListWidgetItem(f"{i+1}. {state['description']}")
            self.redo_list.addItem(item)
        
        # Update counts
        undo_count = self.undo_redo_manager.get_undo_count()
        redo_count = self.undo_redo_manager.get_redo_count()
        
        self.undo_count_label.setText(f"Undo Available: {undo_count}")
        self.redo_count_label.setText(f"Redo Available: {redo_count}")
    
    def perform_undo(self):
        """Perform undo and refresh"""
        if self.undo_redo_manager.undo():
            self.load_history()
    
    def perform_redo(self):
        """Perform redo and refresh"""
        if self.undo_redo_manager.redo():
            self.load_history()
    
    def undo_to_item(self, item):
        """Undo to a specific item in history"""
        index = self.undo_list.row(item)
        
        # Perform multiple undos to reach this state
        for _ in range(index):
            if not self.undo_redo_manager.undo():
                break
        
        self.load_history()
    
    def redo_to_item(self, item):
        """Redo to a specific item in history"""
        index = self.redo_list.row(item)
        
        # Perform multiple redos to reach this state
        for _ in range(index + 1):
            if not self.undo_redo_manager.redo():
                break
        
        self.load_history()
    
    def clear_history(self):
        """Clear all history"""
        reply = QMessageBox.question(
            self,
            "Clear History",
            "Are you sure you want to clear all undo/redo history?\n"
            "This action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.undo_redo_manager.clear()
            self.load_history()
            QMessageBox.information(self, "Success", "History cleared!")
