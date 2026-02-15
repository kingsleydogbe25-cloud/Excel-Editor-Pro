# Excel Editor Pro V3.4 - Cloud Sync & AI Edition

## Overview
Excel Editor Pro is a powerful, user-friendly application for editing Excel (.xlsx, .xls) and CSV files. It combines ease of use with professional features, making data management accessible to everyone. Version 3.4 introduces seamless Cloud Synchronization and powerful AI capabilities.

## Installation

### System Requirements
*   **Operating System:** Windows 7/10/11, macOS 10.13+, or Linux
*   **Python:** 3.7 or higher

### Dependencies
To install all required dependencies, run:

```bash
pip install PyQt5 pandas numpy openpyxl python-docx scikit-learn scipy google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client msal dropbox
```

Or install from the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

## Quick Start Guide

### 1. Creating a New File
1.  Click the **New File** button in the toolbar.
2.  Choose file type (Excel .xlsx or CSV .csv).
3.  Set initial dimensions (rows and columns).
4.  Choose whether to include headers.
5.  Click **OK** to create your file.

### 2. Opening an Existing File
1.  Click the **Open File** button.
2.  Browse to your Excel (.xlsx, .xls) or CSV file.
3.  If the file has multiple sheets, select the one you want to edit.

### 3. Editing Data
*   **Edit cells:** Double-click any cell to edit its content.
*   **Add rows:** Click "Add Row" button and choose position.
*   **Add columns:** Click "Add Column" and specify details.
*   **Delete rows:** Select a row and click "Delete Row".
*   **Reorder rows:** Use "Move Up ↑" and "Move Down ↓" buttons.

### 4. Saving Your Work
*   **Save:** Click "Save" (Ctrl+S) to save changes to the current file.
*   **Save As:** Click "Save As" (Ctrl+Shift+S) to save with a new name or format.
*   **Cloud Upload:** Press Ctrl+U for quick cloud upload.

## Key Features

### ☁️ Cloud Synchronization (New in v3.4)
Seamlessly backup and sync your files with Google Drive, OneDrive, and Dropbox.

*   **Connect Services:** managing your cloud connections via the **Cloud Sync** dialog (Ctrl+Shift+U).
*   **Quick Upload:** Press **Ctrl+U** to instantly upload your current file.
*   **Auto-Sync:** Enable automatic backups on every save in Settings.
*   **Selective Sync:** Choose specific folders to keep in sync automatically.

### 🤖 AI Features (Enhanced)
Leverage local machine learning to analyze and clean your data.

*   **Smart Insights:** Automatic detection of patterns, column types, and anomalies.
*   **Data Cleaning:** Intelligent filling of missing values (KNN imputation) and duplicate detection.
*   **Predictions:** Predict missing values or forecast trends using machine learning.
*   **Formula Assistant:** Get AI-generated Excel formula suggestions.
*   **Access:** Open the AI Assistant with **Ctrl+Shift+A**.

### 🎨 Formatting & Visualization
*   **Rich Formatting:** Custom fonts, colors, borders, and alignment.
*   **Visual Feedback:** Changes appear immediately in the editor.
*   **Number Formats:** Currency, Percentage, Date, Scientific, and more.
*   **Themes:** Switch between Light and Dark modes.

### 🔄 Data Transformation
Powerful tools to clean and modify data (Ctrl+T):
*   **Text:** Change case, trim spaces, find/replace, split text.
*   **Math:** Basic operations, rounding, percentage conversion.
*   **Dates:** Parse text to dates, format dates, date arithmetic.
*   **Cleaning:** Remove duplicates, handle missing values (fill forward/backward/value).

## Keyboard Shortcuts

| Action | Shortcut | Description |
|:-------|:---------|:------------|
| **File Operations** | | |
| New File | `Ctrl+N` | Create a new file |
| Open File | `Ctrl+O` | Open an existing file |
| Save | `Ctrl+S` | Save current file |
| Save As | `Ctrl+Shift+S` | Save as new file |
| Cloud Sync | `Ctrl+Shift+U` | Open Cloud Sync dialog |
| Quick Upload | `Ctrl+U` | Upload to default cloud service |
| **Editing** | | |
| Undo | `Ctrl+Z` | Undo last action |
| Redo | `Ctrl+Y` | Redo last action |
| Add Row | `Ctrl+R` | Insert a new row |
| Add Column | `Ctrl+Shift+C` | Insert a new column |
| Delete Row | `Delete` | Delete selected row |
| **Tools** | | |
| Find/Filter | `Ctrl+F` | Focus filter box |
| Format Columns | `Ctrl+Shift+F` | Open formatting dialog |
| Data Transform | `Ctrl+T` | Open transformation tools |
| AI Assistant | `Ctrl+Shift+A` | Open AI features |
| Quick AI | `Ctrl+Alt+A` | Run quick analysis |
| Statistics | `Ctrl+I` | Show dataset statistics |
| Settings | `Ctrl+,` | Open settings |
| Help | `F1` | Open help dialog |

## Troubleshooting

### Common Issues

*   **"No module named 'openpyxl'":**
    *   Run `pip install openpyxl` to install the required library.
    *   For full features, use the installation command listed above.

*   **File Won't Open:**
    *   Ensure the file is not open in Excel or another program.
    *   Verify the file extension (.csv vs .xlsx).

*   **Cloud Connection Fails:**
    *   Check your internet connection.
    *   Verify your firewall isn't blocking the application.
    *   Try re-authenticating the service in the Cloud Sync tab.

*   **AI Features Unavailable:**
    *   Ensure `scikit-learn` and `scipy` are installed: `pip install scikit-learn scipy`.

### Performance Tips
*   For very large files (>100,000 cells), use **Filters** to work with subsets of data.
*   Disable **Auto-Save** if saving becomes too slow with massive files.
*   Use the 64-bit version of Python for better memory management.

## Developer Configuration
**Important:** To enable robust Cloud Synchronization with OAuth, you must configure your own API credentials.

1.  Open `CloudSync_OAuth.py` in a text editor.
2.  Locate the `OAUTH_CONFIGS` dictionary near the top of the file.
3.  Replace the placeholder values (`YOUR_CLIENT_ID`, etc.) with your actual credentials from Google Cloud Console, Microsoft Azure Portal, or Dropbox App Console.
4.  Set your redirect URI to `http://localhost:8080/callback` in your app settings on the provider's dashboard.

## About
**Version:** 3.4.0 - Cloud Sync Edition
**Release Date:** February 8, 2026
**Technologies:** Python, PyQt5, pandas, openpyxl, scikit-learn

Built for efficiency and professional data management.
