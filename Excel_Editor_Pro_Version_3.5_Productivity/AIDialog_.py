"""
AI Features Dialog - User interface for AI-powered operations
"""

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QTabWidget, QWidget, QTextEdit,
                             QListWidget, QListWidgetItem, QGroupBox,
                             QProgressBar, QComboBox, QCheckBox, QMessageBox,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor
import pandas as pd
from typing import Dict, Any, List

try:
    from AIFeatures_ import (AIManager, AIDataCleaner, AIPredictiveAnalyzer, 
                             AIFormulaAssistant, SKLEARN_AVAILABLE)
except ImportError:
    SKLEARN_AVAILABLE = False


class AIAnalysisThread(QThread):
    """Background thread for AI analysis"""
    finished = pyqtSignal(dict)
    progress = pyqtSignal(str)
    
    def __init__(self, df: pd.DataFrame):
        super().__init__()
        self.df = df
        
    def run(self):
        try:
            self.progress.emit("Initializing AI analyzer...")
            ai_manager = AIManager(self.df)
            
            self.progress.emit("Analyzing data structure...")
            results = ai_manager.get_full_analysis()
            
            self.progress.emit("Analysis complete!")
            self.finished.emit(results)
        except Exception as e:
            self.finished.emit({'error': str(e)})


class AIFeaturesDialog(QDialog):
    """Main dialog for AI features"""
    
    def __init__(self, parent=None, df: pd.DataFrame = None):
        super().__init__(parent)
        self.df = df
        self.parent_editor = parent
        self.ai_manager = AIManager(df) if df is not None else None
        self.analysis_results = None
        
        self.setWindowTitle("🤖 AI Data Assistant")
        self.setMinimumSize(900, 700)
        
        self.init_ui()
        
        # Check if AI is available
        if not SKLEARN_AVAILABLE:
            self.show_installation_guide()
    
    def init_ui(self):
        """Initialize the user interface"""
        layout = QVBoxLayout()
        
        # Header
        header = self.create_header()
        layout.addWidget(header)
        
        # Tab widget for different AI features
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 2px solid #cccccc;
                border-radius: 5px;
            }
            QTabBar::tab {
                padding: 10px 20px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: #FBA002;
                color: white;
                font-weight: bold;
            }
        """)
        
        # Add tabs
        self.tabs.addTab(self.create_insights_tab(), "🔍 Smart Insights")
        self.tabs.addTab(self.create_data_quality_tab(), "✅ Data Quality")
        self.tabs.addTab(self.create_cleaning_tab(), "🧹 Smart Cleaning")
        self.tabs.addTab(self.create_predictions_tab(), "🔮 Predictions")
        self.tabs.addTab(self.create_formulas_tab(), "📐 Formula Assistant")
        
        layout.addWidget(self.tabs)
        
        # Action buttons
        button_layout = QHBoxLayout()
        
        self.analyze_btn = QPushButton("🔍 Analyze Data")
        self.analyze_btn.clicked.connect(self.run_analysis)
        self.analyze_btn.setStyleSheet("""
            QPushButton {
                background: #4CAF50;
                color: white;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background: #45a049;
            }
        """)
        
        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.clicked.connect(self.refresh_data)
        
        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.close)
        
        button_layout.addWidget(self.analyze_btn)
        button_layout.addWidget(self.refresh_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.close_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def create_header(self) -> QWidget:
        """Create header widget"""
        header_widget = QWidget()
        header_layout = QVBoxLayout()
        
        title = QLabel("🤖 AI Data Assistant")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        
        subtitle = QLabel("Intelligent analysis and automated data operations")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #666; font-size: 12px;")
        
        # Status indicator
        self.status_label = QLabel()
        self.update_status_label()
        self.status_label.setAlignment(Qt.AlignCenter)
        
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        header_layout.addWidget(self.status_label)
        
        header_widget.setLayout(header_layout)
        return header_widget
    
    def update_status_label(self):
        """Update AI status label"""
        if SKLEARN_AVAILABLE:
            self.status_label.setText("✓ AI Features Active")
            self.status_label.setStyleSheet("color: #4CAF50; font-weight: bold;")
        else:
            self.status_label.setText("⚠ Limited Mode: Install scikit-learn for full features")
            self.status_label.setStyleSheet("color: #FF9800; font-weight: bold;")
    
    def create_insights_tab(self) -> QWidget:
        """Create insights tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Column type detection
        type_group = QGroupBox("📊 Detected Column Types")
        type_layout = QVBoxLayout()
        
        self.column_types_list = QListWidget()
        self.column_types_list.setStyleSheet("""
            QListWidget {
                font-family: 'Courier New', monospace;
                background: #f5f5f5;
            }
            QListWidget::item {
                padding: 5px;
            }
        """)
        type_layout.addWidget(self.column_types_list)
        
        type_group.setLayout(type_layout)
        layout.addWidget(type_group)
        
        # Key insights
        insights_group = QGroupBox("💡 Key Insights")
        insights_layout = QVBoxLayout()
        
        self.insights_text = QTextEdit()
        self.insights_text.setReadOnly(True)
        self.insights_text.setMaximumHeight(200)
        insights_layout.addWidget(self.insights_text)
        
        insights_group.setLayout(insights_layout)
        layout.addWidget(insights_group)
        
        # Correlations
        corr_group = QGroupBox("🔗 Strong Correlations")
        corr_layout = QVBoxLayout()
        
        self.correlations_list = QListWidget()
        corr_layout.addWidget(self.correlations_list)
        
        corr_group.setLayout(corr_layout)
        layout.addWidget(corr_group)
        
        widget.setLayout(layout)
        return widget
    
    def create_data_quality_tab(self) -> QWidget:
        """Create data quality tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Quality score
        score_group = QGroupBox("📈 Overall Quality Score")
        score_layout = QVBoxLayout()
        
        self.quality_score_label = QLabel("Score: --")
        self.quality_score_label.setFont(QFont("Arial", 24, QFont.Bold))
        self.quality_score_label.setAlignment(Qt.AlignCenter)
        
        self.quality_progress = QProgressBar()
        self.quality_progress.setMinimum(0)
        self.quality_progress.setMaximum(100)
        self.quality_progress.setTextVisible(True)
        self.quality_progress.setFormat("%p%")
        self.quality_progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid grey;
                border-radius: 5px;
                text-align: center;
                height: 30px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
            }
        """)
        
        score_layout.addWidget(self.quality_score_label)
        score_layout.addWidget(self.quality_progress)
        
        score_group.setLayout(score_layout)
        layout.addWidget(score_group)
        
        # Issues list
        issues_group = QGroupBox("⚠️ Detected Issues")
        issues_layout = QVBoxLayout()
        
        self.issues_list = QListWidget()
        issues_layout.addWidget(self.issues_list)
        
        issues_group.setLayout(issues_layout)
        layout.addWidget(issues_group)
        
        # Anomalies
        anomalies_group = QGroupBox("🔍 Anomalies Detected")
        anomalies_layout = QVBoxLayout()
        
        self.anomalies_list = QListWidget()
        anomalies_layout.addWidget(self.anomalies_list)
        
        anomalies_group.setLayout(anomalies_layout)
        layout.addWidget(anomalies_group)
        
        widget.setLayout(layout)
        return widget
    
    def create_cleaning_tab(self) -> QWidget:
        """Create data cleaning tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        info_label = QLabel("Select a column and apply AI-powered cleaning operations")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # Column selection
        col_layout = QHBoxLayout()
        col_layout.addWidget(QLabel("Column:"))
        self.clean_column_combo = QComboBox()
        if self.df is not None:
            self.clean_column_combo.addItems(self.df.columns.tolist())
        col_layout.addWidget(self.clean_column_combo)
        layout.addLayout(col_layout)
        
        # Operations
        ops_group = QGroupBox("🧹 Cleaning Operations")
        ops_layout = QVBoxLayout()
        
        # Missing values
        missing_layout = QHBoxLayout()
        missing_layout.addWidget(QLabel("Fill Missing Values:"))
        self.fill_method_combo = QComboBox()
        self.fill_method_combo.addItems(['Auto', 'Mean', 'Median', 'Mode', 'Forward Fill', 'KNN (Smart)'])
        missing_layout.addWidget(self.fill_method_combo)
        self.fill_missing_btn = QPushButton("Apply")
        self.fill_missing_btn.clicked.connect(self.apply_fill_missing)
        missing_layout.addWidget(self.fill_missing_btn)
        ops_layout.addLayout(missing_layout)
        
        # Remove outliers
        outlier_layout = QHBoxLayout()
        outlier_layout.addWidget(QLabel("Remove Outliers:"))
        self.outlier_method_combo = QComboBox()
        self.outlier_method_combo.addItems(['Z-Score', 'IQR'])
        outlier_layout.addWidget(self.outlier_method_combo)
        self.remove_outliers_btn = QPushButton("Apply")
        self.remove_outliers_btn.clicked.connect(self.apply_remove_outliers)
        outlier_layout.addWidget(self.remove_outliers_btn)
        ops_layout.addLayout(outlier_layout)
        
        # Text standardization
        text_layout = QVBoxLayout()
        text_layout.addWidget(QLabel("Text Standardization:"))
        
        self.text_ops_checks = {}
        text_ops = ['Trim Whitespace', 'Remove Extra Spaces', 'Title Case', 
                   'Lower Case', 'Upper Case', 'Remove Special Characters']
        
        for op in text_ops:
            checkbox = QCheckBox(op)
            self.text_ops_checks[op] = checkbox
            text_layout.addWidget(checkbox)
        
        self.standardize_btn = QPushButton("Apply Text Operations")
        self.standardize_btn.clicked.connect(self.apply_text_standardization)
        text_layout.addWidget(self.standardize_btn)
        
        ops_layout.addLayout(text_layout)
        ops_group.setLayout(ops_layout)
        layout.addWidget(ops_group)
        
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget
    
    def create_predictions_tab(self) -> QWidget:
        """Create predictions tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        info_label = QLabel("Use AI to predict missing values or forecast trends")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # Predict missing values
        predict_group = QGroupBox("🎯 Predict Missing Values")
        predict_layout = QVBoxLayout()
        
        pred_col_layout = QHBoxLayout()
        pred_col_layout.addWidget(QLabel("Target Column:"))
        self.predict_column_combo = QComboBox()
        if self.df is not None:
            numeric_cols = self.df.select_dtypes(include=['number']).columns.tolist()
            self.predict_column_combo.addItems(numeric_cols)
        pred_col_layout.addWidget(self.predict_column_combo)
        
        self.predict_btn = QPushButton("🔮 Predict")
        self.predict_btn.clicked.connect(self.apply_predict_missing)
        pred_col_layout.addWidget(self.predict_btn)
        
        predict_layout.addLayout(pred_col_layout)
        
        self.predict_info_label = QLabel()
        self.predict_info_label.setWordWrap(True)
        self.predict_info_label.setStyleSheet("color: #666; font-style: italic;")
        predict_layout.addWidget(self.predict_info_label)
        
        predict_group.setLayout(predict_layout)
        layout.addWidget(predict_group)
        
        # Forecast trend
        forecast_group = QGroupBox("📈 Forecast Trend")
        forecast_layout = QVBoxLayout()
        
        fore_col_layout = QHBoxLayout()
        fore_col_layout.addWidget(QLabel("Column:"))
        self.forecast_column_combo = QComboBox()
        if self.df is not None:
            numeric_cols = self.df.select_dtypes(include=['number']).columns.tolist()
            self.forecast_column_combo.addItems(numeric_cols)
        fore_col_layout.addWidget(self.forecast_column_combo)
        
        fore_col_layout.addWidget(QLabel("Periods:"))
        self.forecast_periods_combo = QComboBox()
        self.forecast_periods_combo.addItems(['3', '5', '10'])
        fore_col_layout.addWidget(self.forecast_periods_combo)
        
        self.forecast_btn = QPushButton("📊 Forecast")
        self.forecast_btn.clicked.connect(self.apply_forecast)
        fore_col_layout.addWidget(self.forecast_btn)
        
        forecast_layout.addLayout(fore_col_layout)
        
        self.forecast_results = QTextEdit()
        self.forecast_results.setReadOnly(True)
        self.forecast_results.setMaximumHeight(150)
        forecast_layout.addWidget(self.forecast_results)
        
        forecast_group.setLayout(forecast_layout)
        layout.addWidget(forecast_group)
        
        # Clustering
        cluster_group = QGroupBox("🎨 Smart Clustering")
        cluster_layout = QVBoxLayout()
        
        cluster_control_layout = QHBoxLayout()
        cluster_control_layout.addWidget(QLabel("Number of Clusters:"))
        self.cluster_count_combo = QComboBox()
        self.cluster_count_combo.addItems(['2', '3', '4', '5'])
        cluster_control_layout.addWidget(self.cluster_count_combo)
        
        self.cluster_btn = QPushButton("🎨 Create Clusters")
        self.cluster_btn.clicked.connect(self.apply_clustering)
        cluster_control_layout.addWidget(self.cluster_btn)
        
        cluster_layout.addLayout(cluster_control_layout)
        
        self.cluster_info_label = QLabel()
        self.cluster_info_label.setWordWrap(True)
        self.cluster_info_label.setStyleSheet("color: #666; font-style: italic;")
        cluster_layout.addWidget(self.cluster_info_label)
        
        cluster_group.setLayout(cluster_layout)
        layout.addWidget(cluster_group)
        
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget
    
    def create_formulas_tab(self) -> QWidget:
        """Create formula assistant tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        info_label = QLabel("Get AI-powered formula suggestions based on your data")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # Formula suggestions
        suggestions_group = QGroupBox("💡 Suggested Formulas")
        suggestions_layout = QVBoxLayout()
        
        self.formula_suggestions_list = QListWidget()
        self.formula_suggestions_list.itemClicked.connect(self.on_formula_selected)
        suggestions_layout.addWidget(self.formula_suggestions_list)
        
        self.generate_suggestions_btn = QPushButton("🔄 Generate Suggestions")
        self.generate_suggestions_btn.clicked.connect(self.generate_formula_suggestions)
        suggestions_layout.addWidget(self.generate_suggestions_btn)
        
        suggestions_group.setLayout(suggestions_layout)
        layout.addWidget(suggestions_group)
        
        # Formula details
        details_group = QGroupBox("📋 Formula Details")
        details_layout = QVBoxLayout()
        
        self.formula_details_text = QTextEdit()
        self.formula_details_text.setReadOnly(True)
        self.formula_details_text.setMaximumHeight(150)
        details_layout.addWidget(self.formula_details_text)
        
        details_group.setLayout(details_layout)
        layout.addWidget(details_group)
        
        widget.setLayout(layout)
        return widget
    
    def run_analysis(self):
        """Run full AI analysis"""
        if self.df is None:
            QMessageBox.warning(self, "No Data", "Please load data first")
            return
        
        if not SKLEARN_AVAILABLE:
            QMessageBox.warning(self, "AI Not Available", 
                              "Install scikit-learn: pip install scikit-learn scipy")
            return
        
        # Show progress
        self.analyze_btn.setEnabled(False)
        self.analyze_btn.setText("Analyzing...")
        
        # Run analysis in background thread
        self.analysis_thread = AIAnalysisThread(self.df)
        self.analysis_thread.finished.connect(self.on_analysis_complete)
        self.analysis_thread.start()
    
    def on_analysis_complete(self, results: Dict[str, Any]):
        """Handle analysis completion"""
        self.analyze_btn.setEnabled(True)
        self.analyze_btn.setText("🔍 Analyze Data")
        
        if 'error' in results:
            QMessageBox.critical(self, "Analysis Error", f"Error: {results['error']}")
            return
        
        self.analysis_results = results
        self.display_results(results)
    
    def display_results(self, results: Dict[str, Any]):
        """Display analysis results"""
        # Column types
        self.column_types_list.clear()
        if 'column_types' in results:
            for col, col_type in results['column_types'].items():
                item = QListWidgetItem(f"  {col:<30} → {col_type}")
                self.column_types_list.addItem(item)
        
        # Insights
        if 'insights' in results:
            insights_text = "\n\n".join(f"• {insight}" for insight in results['insights'])
            self.insights_text.setText(insights_text)
        
        # Correlations
        self.correlations_list.clear()
        if 'correlations' in results:
            for col1, col2, corr in results['correlations']:
                item_text = f"  {col1} ↔ {col2}: {corr:.3f}"
                item = QListWidgetItem(item_text)
                if abs(corr) > 0.9:
                    item.setForeground(QColor('red'))
                self.correlations_list.addItem(item)
        
        # Data quality
        if 'data_quality' in results:
            quality = results['data_quality']
            score = quality.get('score', 0)
            
            self.quality_score_label.setText(f"Score: {score:.1f}/100")
            self.quality_progress.setValue(int(score))
            
            # Color code the score
            if score >= 90:
                color = "#4CAF50"  # Green
            elif score >= 70:
                color = "#FF9800"  # Orange
            else:
                color = "#F44336"  # Red
            
            self.quality_progress.setStyleSheet(f"""
                QProgressBar {{
                    border: 2px solid grey;
                    border-radius: 5px;
                    text-align: center;
                    height: 30px;
                }}
                QProgressBar::chunk {{
                    background-color: {color};
                }}
            """)
            
            # Issues
            self.issues_list.clear()
            for issue in quality.get('issues', []):
                self.issues_list.addItem(f"  ⚠ {issue}")
        
        # Anomalies
        self.anomalies_list.clear()
        if 'anomalies' in results:
            for col, indices in results['anomalies'].items():
                item_text = f"  {col}: {len(indices)} anomalies at rows {indices[:5]}"
                self.anomalies_list.addItem(item_text)
    
    def apply_fill_missing(self):
        """Apply missing value filling"""
        if not self.check_sklearn():
            return
        
        column = self.clean_column_combo.currentText()
        method_map = {
            'Auto': 'auto',
            'Mean': 'mean',
            'Median': 'median',
            'Mode': 'mode',
            'Forward Fill': 'forward',
            'KNN (Smart)': 'knn'
        }
        method = method_map[self.fill_method_combo.currentText()]
        
        try:
            filled_series = AIDataCleaner.smart_fill_missing(self.df, column, method)
            self.df[column] = filled_series
            
            if self.parent_editor:
                self.parent_editor.df = self.df
                self.parent_editor.display_dataframe()
                self.parent_editor.is_modified = True
            
            QMessageBox.information(self, "Success", f"Filled missing values in '{column}'")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to fill missing values: {str(e)}")
    
    def apply_remove_outliers(self):
        """Apply outlier removal"""
        if not self.check_sklearn():
            return
        
        column = self.clean_column_combo.currentText()
        method_map = {'Z-Score': 'zscore', 'IQR': 'iqr'}
        method = method_map[self.outlier_method_combo.currentText()]
        
        try:
            cleaned_series = AIDataCleaner.remove_outliers(self.df[column], method)
            outliers_removed = cleaned_series.isna().sum() - self.df[column].isna().sum()
            self.df[column] = cleaned_series
            
            if self.parent_editor:
                self.parent_editor.df = self.df
                self.parent_editor.display_dataframe()
                self.parent_editor.is_modified = True
            
            QMessageBox.information(self, "Success", 
                                  f"Removed {outliers_removed} outliers from '{column}'")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to remove outliers: {str(e)}")
    
    def apply_text_standardization(self):
        """Apply text standardization operations"""
        column = self.clean_column_combo.currentText()
        
        operations = []
        op_map = {
            'Trim Whitespace': 'trim',
            'Remove Extra Spaces': 'remove_extra_spaces',
            'Title Case': 'title',
            'Lower Case': 'lower',
            'Upper Case': 'upper',
            'Remove Special Characters': 'remove_special'
        }
        
        for op_name, op_key in op_map.items():
            if self.text_ops_checks[op_name].isChecked():
                operations.append(op_key)
        
        if not operations:
            QMessageBox.warning(self, "No Operations", "Please select at least one operation")
            return
        
        try:
            standardized_series = AIDataCleaner.standardize_text(self.df[column], operations)
            self.df[column] = standardized_series
            
            if self.parent_editor:
                self.parent_editor.df = self.df
                self.parent_editor.display_dataframe()
                self.parent_editor.is_modified = True
            
            QMessageBox.information(self, "Success", f"Standardized text in '{column}'")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to standardize text: {str(e)}")
    
    def apply_predict_missing(self):
        """Apply predictive missing value filling"""
        if not self.check_sklearn():
            return
        
        column = self.predict_column_combo.currentText()
        missing_count = self.df[column].isna().sum()
        
        if missing_count == 0:
            QMessageBox.information(self, "No Missing Values", 
                                  f"Column '{column}' has no missing values")
            return
        
        try:
            predicted_series = AIPredictiveAnalyzer.predict_missing_values(self.df, column)
            self.df[column] = predicted_series
            
            if self.parent_editor:
                self.parent_editor.df = self.df
                self.parent_editor.display_dataframe()
                self.parent_editor.is_modified = True
            
            self.predict_info_label.setText(
                f"✓ Predicted {missing_count} missing values in '{column}' using ML"
            )
            QMessageBox.information(self, "Success", 
                                  f"Predicted {missing_count} missing values")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Prediction failed: {str(e)}")
    
    def apply_forecast(self):
        """Apply trend forecasting"""
        if not self.check_sklearn():
            return
        
        column = self.forecast_column_combo.currentText()
        periods = int(self.forecast_periods_combo.currentText())
        
        try:
            forecasts = AIPredictiveAnalyzer.forecast_trend(self.df[column], periods)
            
            if not forecasts:
                QMessageBox.warning(self, "Insufficient Data", 
                                  "Not enough data for forecasting")
                return
            
            results_text = f"Forecast for '{column}' (next {periods} periods):\n\n"
            current_max = self.df[column].max()
            
            for i, value in enumerate(forecasts, 1):
                trend = "↑" if value > current_max else "↓"
                results_text += f"Period {i}: {value:.2f} {trend}\n"
                current_max = value
            
            self.forecast_results.setText(results_text)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Forecasting failed: {str(e)}")
    
    def apply_clustering(self):
        """Apply k-means clustering"""
        if not self.check_sklearn():
            return
        
        n_clusters = int(self.cluster_count_combo.currentText())
        
        try:
            clusters = AIPredictiveAnalyzer.cluster_data(self.df, n_clusters)
            
            if len(clusters) == 0:
                QMessageBox.warning(self, "No Numeric Data", 
                                  "Need numeric columns for clustering")
                return
            
            # Add cluster column
            cluster_col_name = 'AI_Cluster'
            self.df[cluster_col_name] = clusters
            
            if self.parent_editor:
                self.parent_editor.df = self.df
                self.parent_editor.display_dataframe()
                self.parent_editor.is_modified = True
            
            # Show distribution
            cluster_counts = pd.Series(clusters).value_counts().sort_index()
            info_text = f"Created {n_clusters} clusters:\n\n"
            for cluster_id, count in cluster_counts.items():
                info_text += f"Cluster {cluster_id}: {count} rows\n"
            
            self.cluster_info_label.setText(info_text)
            
            QMessageBox.information(self, "Success", 
                                  f"Created '{cluster_col_name}' column with {n_clusters} clusters")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Clustering failed: {str(e)}")
    
    def generate_formula_suggestions(self):
        """Generate formula suggestions"""
        if self.analysis_results is None or 'column_types' not in self.analysis_results:
            QMessageBox.information(self, "Run Analysis First", 
                                  "Please run analysis first to get suggestions")
            return
        
        column_types = self.analysis_results['column_types']
        suggestions = AIFormulaAssistant.suggest_formulas(self.df, column_types)
        
        self.formula_suggestions_list.clear()
        
        for suggestion in suggestions:
            item = QListWidgetItem(f"📐 {suggestion['name']}")
            item.setData(Qt.UserRole, suggestion)
            self.formula_suggestions_list.addItem(item)
    
    def on_formula_selected(self, item: QListWidgetItem):
        """Handle formula selection"""
        suggestion = item.data(Qt.UserRole)
        
        details = f"Formula: {suggestion['formula']}\n\n"
        details += f"Description: {suggestion['description']}\n\n"
        details += f"Explanation: {AIFormulaAssistant.explain_formula(suggestion['formula'])}"
        
        self.formula_details_text.setText(details)
    
    def refresh_data(self):
        """Refresh data from parent"""
        if self.parent_editor and hasattr(self.parent_editor, 'df'):
            self.df = self.parent_editor.df
            if self.df is not None:
                self.ai_manager = AIManager(self.df)
                self.analysis_results = None
                
                # Update combo boxes
                columns = self.df.columns.tolist()
                self.clean_column_combo.clear()
                self.clean_column_combo.addItems(columns)
                
                numeric_cols = self.df.select_dtypes(include=['number']).columns.tolist()
                self.predict_column_combo.clear()
                self.predict_column_combo.addItems(numeric_cols)
                self.forecast_column_combo.clear()
                self.forecast_column_combo.addItems(numeric_cols)
                
                QMessageBox.information(self, "Refreshed", "Data refreshed successfully")
    
    def check_sklearn(self) -> bool:
        """Check if sklearn is available"""
        if not SKLEARN_AVAILABLE:
            QMessageBox.warning(self, "Feature Not Available",
                              "This feature requires scikit-learn.\n\n"
                              "Install with: pip install scikit-learn scipy")
            return False
        return True
    
    def show_installation_guide(self):
        """Show installation guide for AI features"""
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("AI Features Setup")
        msg.setText("To enable full AI features, install the following packages:")
        msg.setInformativeText(
            "pip install scikit-learn scipy\n\n"
            "These packages provide:\n"
            "• Smart data analysis\n"
            "• Anomaly detection\n"
            "• Predictive analytics\n"
            "• Intelligent clustering\n"
            "• Advanced imputation"
        )
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()
