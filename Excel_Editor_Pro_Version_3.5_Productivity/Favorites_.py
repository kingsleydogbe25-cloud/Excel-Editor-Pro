"""
Favorites/Bookmarks Manager
Allows users to mark and quickly access frequently used files
"""

import os
import json
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                             QListWidget, QLabel, QMessageBox, QFileDialog,
                             QListWidgetItem, QInputDialog)
from PyQt5.QtCore import QSettings, Qt
from PyQt5.QtGui import QIcon
from datetime import datetime


class FavoritesManager:
    """Manages favorite/bookmarked files"""
    
    def __init__(self, parent):
        self.parent = parent
        self.settings = QSettings("DataTools", "ExcelEditor")
        
        # Load favorites from settings
        self.favorites = self.load_favorites()
    
    def load_favorites(self):
        """Load favorites from settings"""
        try:
            favorites_json = self.settings.value("favorites", "[]")
            if isinstance(favorites_json, str):
                favorites = json.loads(favorites_json)
            else:
                favorites = []
            return favorites
        except:
            return []
    
    def save_favorites(self):
        """Save favorites to settings"""
        try:
            favorites_json = json.dumps(self.favorites)
            self.settings.setValue("favorites", favorites_json)
        except Exception as e:
            print(f"Error saving favorites: {e}")
    
    def add_favorite(self, file_path, description=None):
        """Add a file to favorites"""
        try:
            # Check if already in favorites
            for fav in self.favorites:
                if fav['path'] == file_path:
                    QMessageBox.information(
                        self.parent,
                        "Already Favorited",
                        "This file is already in your favorites!"
                    )
                    return False
            
            # Get description if not provided
            if description is None:
                description, ok = QInputDialog.getText(
                    self.parent,
                    "Add to Favorites",
                    "Enter a description (optional):",
                    text=os.path.basename(file_path)
                )
                if not ok:
                    return False
            
            # Add to favorites
            favorite = {
                'path': file_path,
                'description': description or os.path.basename(file_path),
                'added': datetime.now().isoformat(),
                'last_opened': None
            }
            
            self.favorites.append(favorite)
            self.save_favorites()
            
            self.parent.statusBar().showMessage("Added to favorites", 2000)
            return True
            
        except Exception as e:
            QMessageBox.critical(
                self.parent,
                "Error",
                f"Failed to add favorite:\n{str(e)}"
            )
            return False
    
    def remove_favorite(self, file_path):
        """Remove a file from favorites"""
        try:
            self.favorites = [f for f in self.favorites if f['path'] != file_path]
            self.save_favorites()
            return True
        except Exception as e:
            print(f"Error removing favorite: {e}")
            return False
    
    def update_last_opened(self, file_path):
        """Update last opened time for a favorite"""
        try:
            for fav in self.favorites:
                if fav['path'] == file_path:
                    fav['last_opened'] = datetime.now().isoformat()
                    self.save_favorites()
                    break
        except Exception as e:
            print(f"Error updating last opened: {e}")
    
    def is_favorite(self, file_path):
        """Check if a file is in favorites"""
        return any(fav['path'] == file_path for fav in self.favorites)
    
    def get_favorites(self):
        """Get all favorites"""
        return self.favorites.copy()
    
    def get_sorted_favorites(self, sort_by='added'):
        """Get favorites sorted by specified criterion"""
        favorites = self.favorites.copy()
        
        if sort_by == 'name':
            favorites.sort(key=lambda x: x['description'].lower())
        elif sort_by == 'last_opened':
            favorites.sort(key=lambda x: x.get('last_opened') or '', reverse=True)
        elif sort_by == 'path':
            favorites.sort(key=lambda x: x['path'].lower())
        else:  # added
            favorites.sort(key=lambda x: x.get('added', ''), reverse=True)
        
        return favorites


class FavoritesDialog(QDialog):
    """Dialog to manage favorite files"""
    
    def __init__(self, parent, favorites_manager):
        super().__init__(parent)
        self.parent = parent
        self.favorites_manager = favorites_manager
        self.init_ui()
        self.load_favorites()
    
    def init_ui(self):
        self.setWindowTitle("Favorite Files")
        self.setGeometry(200, 200, 600, 500)
        
        layout = QVBoxLayout()
        
        # Header
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("⭐ Your Favorite Files"))
        header_layout.addStretch()
        
        # Sort options
        header_layout.addWidget(QLabel("Sort by:"))
        from PyQt5.QtWidgets import QComboBox
        self.sort_combo = QComboBox()
        self.sort_combo.addItems(['Recently Added', 'Name', 'Last Opened', 'Path'])
        self.sort_combo.currentTextChanged.connect(self.on_sort_changed)
        header_layout.addWidget(self.sort_combo)
        
        layout.addLayout(header_layout)
        
        # Favorites list
        self.favorites_list = QListWidget()
        self.favorites_list.itemDoubleClicked.connect(self.open_favorite)
        layout.addWidget(self.favorites_list)
        
        # Info label
        self.info_label = QLabel()
        layout.addWidget(self.info_label)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        open_btn = QPushButton("Open Selected")
        open_btn.clicked.connect(self.open_favorite)
        button_layout.addWidget(open_btn)
        
        add_btn = QPushButton("Add Current File")
        add_btn.clicked.connect(self.add_current_file)
        button_layout.addWidget(add_btn)
        
        browse_btn = QPushButton("Add from Browse...")
        browse_btn.clicked.connect(self.browse_and_add)
        button_layout.addWidget(browse_btn)
        
        edit_btn = QPushButton("Edit Description")
        edit_btn.clicked.connect(self.edit_description)
        button_layout.addWidget(edit_btn)
        
        remove_btn = QPushButton("Remove")
        remove_btn.clicked.connect(self.remove_favorite)
        button_layout.addWidget(remove_btn)
        
        button_layout.addStretch()
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def on_sort_changed(self, text):
        """Handle sort option change"""
        sort_map = {
            'Recently Added': 'added',
            'Name': 'name',
            'Last Opened': 'last_opened',
            'Path': 'path'
        }
        self.current_sort = sort_map.get(text, 'added')
        self.load_favorites()
    
    def load_favorites(self):
        """Load favorites into list"""
        self.favorites_list.clear()
        
        sort_by = getattr(self, 'current_sort', 'added')
        favorites = self.favorites_manager.get_sorted_favorites(sort_by)
        
        if not favorites:
            self.favorites_list.addItem("No favorites yet. Add files to get started!")
            self.info_label.setText("Total favorites: 0")
            return
        
        for fav in favorites:
            # Check if file exists
            exists = os.path.exists(fav['path'])
            
            # Format display text
            display_text = fav['description']
            if not exists:
                display_text += " [FILE NOT FOUND]"
            
            # Add item
            item = QListWidgetItem(display_text)
            item.setData(Qt.UserRole, fav)
            
            # Color code based on existence
            if not exists:
                from PyQt5.QtGui import QBrush, QColor
                item.setForeground(QBrush(QColor(150, 150, 150)))
            
            self.favorites_list.addItem(item)
        
        self.info_label.setText(f"Total favorites: {len(favorites)}")
    
    def open_favorite(self):
        """Open selected favorite"""
        current_item = self.favorites_list.currentItem()
        if not current_item:
            return
        
        fav = current_item.data(Qt.UserRole)
        if not fav:
            return
        
        file_path = fav['path']
        
        # Check if file exists
        if not os.path.exists(file_path):
            QMessageBox.warning(
                self,
                "File Not Found",
                f"The file no longer exists:\n{file_path}\n\n"
                "Would you like to remove it from favorites?"
            )
            return
        
        # Check if current file has unsaved changes
        if self.parent.is_modified:
            reply = QMessageBox.question(
                self,
                "Unsaved Changes",
                "Current file has unsaved changes. Do you want to save before opening?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
            )
            
            if reply == QMessageBox.Save:
                self.parent.save_file()
            elif reply == QMessageBox.Cancel:
                return
        
        # Open the file
        try:
            import pandas as pd
            
            self.parent.statusBar().showMessage(f"Loading {file_path}...")
            
            if file_path.endswith('.csv'):
                self.parent.df = pd.read_csv(file_path)
            else:
                self.parent.df = pd.read_excel(file_path)
            
            self.parent.filtered_df = None
            self.parent.current_file_path = file_path
            self.parent.update_table()
            self.parent.is_modified = False
            
            # Update window title
            self.parent.setWindowTitle(f"Excel Editor Pro - {os.path.basename(file_path)}")
            
            # Update last opened
            self.favorites_manager.update_last_opened(file_path)
            
            # Add to recent files
            if hasattr(self.parent, 'add_to_recent_files'):
                self.parent.add_to_recent_files(file_path)
            
            self.parent.statusBar().showMessage(f"Loaded {os.path.basename(file_path)}", 3000)
            
            self.close()
            
        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to open file:\n{str(e)}"
            )
    
    def add_current_file(self):
        """Add current file to favorites"""
        if not self.parent.current_file_path:
            QMessageBox.information(
                self,
                "No File",
                "No file is currently open."
            )
            return
        
        if self.favorites_manager.add_favorite(self.parent.current_file_path):
            self.load_favorites()
    
    def browse_and_add(self):
        """Browse for a file to add to favorites"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select File to Add to Favorites",
            "",
            "Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*.*)"
        )
        
        if file_path:
            if self.favorites_manager.add_favorite(file_path):
                self.load_favorites()
    
    def edit_description(self):
        """Edit description of selected favorite"""
        current_item = self.favorites_list.currentItem()
        if not current_item:
            return
        
        fav = current_item.data(Qt.UserRole)
        if not fav:
            return
        
        new_desc, ok = QInputDialog.getText(
            self,
            "Edit Description",
            "Enter new description:",
            text=fav['description']
        )
        
        if ok and new_desc:
            # Update description
            for f in self.favorites_manager.favorites:
                if f['path'] == fav['path']:
                    f['description'] = new_desc
                    break
            
            self.favorites_manager.save_favorites()
            self.load_favorites()
    
    def remove_favorite(self):
        """Remove selected favorite"""
        current_item = self.favorites_list.currentItem()
        if not current_item:
            return
        
        fav = current_item.data(Qt.UserRole)
        if not fav:
            return
        
        reply = QMessageBox.question(
            self,
            "Remove Favorite",
            f"Remove '{fav['description']}' from favorites?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            if self.favorites_manager.remove_favorite(fav['path']):
                self.load_favorites()
