"""
AI Features Module for Excel Editor Pro
Uses local ML models for intelligent data operations
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional, Any
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Import ML libraries
try:
    from sklearn.ensemble import IsolationForest, RandomForestRegressor
    from sklearn.preprocessing import StandardScaler, LabelEncoder
    from sklearn.impute import SimpleImputer, KNNImputer
    from sklearn.cluster import KMeans
    from sklearn.linear_model import LinearRegression
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False
    print("scikit-learn not available - ML features limited")

try:
    from scipy import stats
    SCIPY_AVAILABLE = True
except ImportError:
    SCIPY_AVAILABLE = False


class AIDataAnalyzer:
    """AI-powered data analysis and insights"""
    
    def __init__(self, df: pd.DataFrame):
        self.df = df
        self.insights = []
        
    def analyze(self) -> Dict[str, Any]:
        """Perform comprehensive data analysis"""
        results = {
            'column_types': self.detect_column_types(),
            'data_quality': self.assess_data_quality(),
            'insights': self.generate_insights(),
            'anomalies': self.detect_anomalies(),
            'correlations': self.find_correlations()
        }
        return results
    
    def detect_column_types(self) -> Dict[str, str]:
        """Intelligently detect semantic column types"""
        type_map = {}
        
        for col in self.df.columns:
            col_type = self._classify_column(col, self.df[col])
            type_map[col] = col_type
            
        return type_map
    
    def _classify_column(self, col_name: str, series: pd.Series) -> str:
        """Classify a single column's semantic type"""
        # Skip if too many nulls
        if series.isna().sum() / len(series) > 0.9:
            return "mostly_empty"
        
        # Check for numeric
        if pd.api.types.is_numeric_dtype(series):
            if series.min() >= 0 and series.max() <= 1:
                return "percentage/probability"
            elif self._is_currency(series):
                return "currency"
            elif self._is_year(series):
                return "year"
            elif series.nunique() < 10 and len(series) > 20:
                return "categorical_numeric"
            else:
                return "numeric"
        
        # Check for datetime
        if pd.api.types.is_datetime64_any_dtype(series):
            return "datetime"
        
        # Analyze text patterns
        sample = series.dropna().astype(str).head(100)
        
        if len(sample) == 0:
            return "empty"
        
        # Email detection
        if sample.str.contains(r'^[\w\.-]+@[\w\.-]+\.\w+$', regex=True).mean() > 0.8:
            return "email"
        
        # Phone detection
        if sample.str.contains(r'[\d\-\(\)\+\s]{10,}', regex=True).mean() > 0.7:
            return "phone"
        
        # URL detection
        if sample.str.contains(r'https?://', regex=True).mean() > 0.7:
            return "url"
        
        # Currency text detection
        if sample.str.contains(r'[\$£€¥]\s*\d', regex=True).mean() > 0.7:
            return "currency_text"
        
        # ID detection (alphanumeric with consistent pattern)
        if sample.str.match(r'^[A-Z0-9\-]{5,}$').mean() > 0.8:
            return "identifier"
        
        # Name detection (contains spaces, title case)
        if col_name.lower() in ['name', 'full_name', 'customer', 'client']:
            return "person_name"
        
        # Address detection
        if col_name.lower() in ['address', 'location', 'street']:
            return "address"
        
        # Category detection (low unique count)
        unique_ratio = series.nunique() / len(series)
        if unique_ratio < 0.05 and series.nunique() < 50:
            return "category"
        
        # Default to text
        return "text"
    
    def _is_currency(self, series: pd.Series) -> bool:
        """Check if numeric series represents currency"""
        # Check if values have typical currency patterns (2 decimals often)
        if not pd.api.types.is_numeric_dtype(series):
            return False
        
        sample = series.dropna().head(100)
        if len(sample) == 0:
            return False
        
        # Check decimal places
        has_cents = (sample % 1 != 0).mean() > 0.5
        return has_cents
    
    def _is_year(self, series: pd.Series) -> bool:
        """Check if numeric series represents years"""
        if not pd.api.types.is_numeric_dtype(series):
            return False
        
        sample = series.dropna()
        if len(sample) == 0:
            return False
        
        return (sample >= 1900).all() and (sample <= 2100).all() and (sample % 1 == 0).all()
    
    def assess_data_quality(self) -> Dict[str, Any]:
        """Assess overall data quality"""
        total_cells = self.df.shape[0] * self.df.shape[1]
        missing_cells = self.df.isna().sum().sum()
        
        quality_score = 100 * (1 - missing_cells / total_cells)
        
        issues = []
        
        # Check for missing data
        for col in self.df.columns:
            missing_pct = self.df[col].isna().sum() / len(self.df) * 100
            if missing_pct > 50:
                issues.append(f"High missing data in '{col}': {missing_pct:.1f}%")
            elif missing_pct > 20:
                issues.append(f"Moderate missing data in '{col}': {missing_pct:.1f}%")
        
        # Check for duplicates
        dup_count = self.df.duplicated().sum()
        if dup_count > 0:
            issues.append(f"Found {dup_count} duplicate rows")
        
        # Check for outliers in numeric columns
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            outliers = self._detect_outliers_zscore(self.df[col])
            if outliers > 0:
                issues.append(f"Found {outliers} outliers in '{col}'")
        
        return {
            'score': quality_score,
            'total_cells': total_cells,
            'missing_cells': missing_cells,
            'missing_percentage': (missing_cells / total_cells * 100),
            'issues': issues
        }
    
    def _detect_outliers_zscore(self, series: pd.Series, threshold: float = 3) -> int:
        """Detect outliers using z-score method"""
        if not SCIPY_AVAILABLE:
            return 0
        
        clean_series = series.dropna()
        if len(clean_series) < 10:
            return 0
        
        z_scores = np.abs(stats.zscore(clean_series))
        return (z_scores > threshold).sum()
    
    def generate_insights(self) -> List[str]:
        """Generate intelligent insights about the data"""
        insights = []
        
        # Basic statistics
        insights.append(f"Dataset contains {len(self.df):,} rows and {len(self.df.columns)} columns")
        
        # Numeric insights
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            insights.append(f"Found {len(numeric_cols)} numeric columns for analysis")
            
            # Find highly variable columns
            for col in numeric_cols:
                cv = self.df[col].std() / self.df[col].mean() if self.df[col].mean() != 0 else 0
                if cv > 1:
                    insights.append(f"'{col}' shows high variability (CV: {cv:.2f})")
        
        # Categorical insights
        obj_cols = self.df.select_dtypes(include=['object']).columns
        for col in obj_cols:
            unique_count = self.df[col].nunique()
            if unique_count < 10:
                insights.append(f"'{col}' has {unique_count} categories")
        
        # Missing data patterns
        missing_cols = [col for col in self.df.columns if self.df[col].isna().sum() > 0]
        if missing_cols:
            insights.append(f"{len(missing_cols)} columns have missing values")
        
        # Temporal patterns
        date_cols = self.df.select_dtypes(include=['datetime64']).columns
        if len(date_cols) > 0:
            for col in date_cols:
                date_range = self.df[col].max() - self.df[col].min()
                insights.append(f"'{col}' spans {date_range.days} days")
        
        return insights
    
    def detect_anomalies(self) -> Dict[str, List[int]]:
        """Detect anomalies in numeric columns"""
        if not SKLEARN_AVAILABLE:
            return {}
        
        anomalies = {}
        numeric_cols = self.df.select_dtypes(include=[np.number]).columns
        
        for col in numeric_cols:
            clean_data = self.df[col].dropna()
            if len(clean_data) < 10:
                continue
            
            # Use Isolation Forest
            try:
                iso_forest = IsolationForest(contamination=0.1, random_state=42)
                predictions = iso_forest.fit_predict(clean_data.values.reshape(-1, 1))
                anomaly_indices = clean_data.index[predictions == -1].tolist()
                
                if anomaly_indices:
                    anomalies[col] = anomaly_indices[:10]  # Limit to first 10
            except:
                continue
        
        return anomalies
    
    def find_correlations(self) -> List[Tuple[str, str, float]]:
        """Find significant correlations between numeric columns"""
        numeric_df = self.df.select_dtypes(include=[np.number])
        
        if numeric_df.shape[1] < 2:
            return []
        
        corr_matrix = numeric_df.corr()
        correlations = []
        
        for i in range(len(corr_matrix.columns)):
            for j in range(i + 1, len(corr_matrix.columns)):
                col1 = corr_matrix.columns[i]
                col2 = corr_matrix.columns[j]
                corr_value = corr_matrix.iloc[i, j]
                
                if abs(corr_value) > 0.7 and not np.isnan(corr_value):
                    correlations.append((col1, col2, corr_value))
        
        # Sort by absolute correlation value
        correlations.sort(key=lambda x: abs(x[2]), reverse=True)
        return correlations[:5]  # Top 5


class AIDataCleaner:
    """AI-powered data cleaning operations"""
    
    @staticmethod
    def smart_fill_missing(df: pd.DataFrame, column: str, method: str = 'auto') -> pd.Series:
        """Intelligently fill missing values"""
        series = df[column].copy()
        
        if method == 'auto':
            # Decide based on data type
            if pd.api.types.is_numeric_dtype(series):
                method = 'knn' if SKLEARN_AVAILABLE else 'median'
            else:
                method = 'mode'
        
        if method == 'median':
            return series.fillna(series.median())
        
        elif method == 'mean':
            return series.fillna(series.mean())
        
        elif method == 'mode':
            mode_val = series.mode()
            if len(mode_val) > 0:
                return series.fillna(mode_val[0])
            return series
        
        elif method == 'forward':
            return series.fillna(method='ffill')
        
        elif method == 'backward':
            return series.fillna(method='bfill')
        
        elif method == 'knn' and SKLEARN_AVAILABLE:
            return AIDataCleaner._knn_impute(df, column)
        
        return series
    
    @staticmethod
    def _knn_impute(df: pd.DataFrame, column: str, n_neighbors: int = 5) -> pd.Series:
        """Use KNN imputation for missing values"""
        if not SKLEARN_AVAILABLE:
            return df[column]
        
        numeric_df = df.select_dtypes(include=[np.number])
        
        if column not in numeric_df.columns:
            return df[column]
        
        try:
            imputer = KNNImputer(n_neighbors=n_neighbors)
            imputed_data = imputer.fit_transform(numeric_df)
            col_idx = numeric_df.columns.get_loc(column)
            result = df[column].copy()
            result[result.isna()] = imputed_data[result.isna(), col_idx]
            return result
        except:
            return df[column].fillna(df[column].median())
    
    @staticmethod
    def remove_outliers(series: pd.Series, method: str = 'zscore', threshold: float = 3) -> pd.Series:
        """Remove outliers from numeric series"""
        if not pd.api.types.is_numeric_dtype(series):
            return series
        
        result = series.copy()
        
        if method == 'zscore' and SCIPY_AVAILABLE:
            z_scores = np.abs(stats.zscore(series.dropna()))
            mask = z_scores < threshold
            result.loc[~mask] = np.nan
        
        elif method == 'iqr':
            Q1 = series.quantile(0.25)
            Q3 = series.quantile(0.75)
            IQR = Q3 - Q1
            lower = Q1 - 1.5 * IQR
            upper = Q3 + 1.5 * IQR
            result = series.where((series >= lower) & (series <= upper), np.nan)
        
        return result
    
    @staticmethod
    def standardize_text(series: pd.Series, operations: List[str]) -> pd.Series:
        """Standardize text with multiple operations"""
        result = series.copy()
        
        for op in operations:
            if op == 'trim':
                result = result.str.strip()
            elif op == 'lower':
                result = result.str.lower()
            elif op == 'upper':
                result = result.str.upper()
            elif op == 'title':
                result = result.str.title()
            elif op == 'remove_extra_spaces':
                result = result.str.replace(r'\s+', ' ', regex=True)
            elif op == 'remove_special':
                result = result.str.replace(r'[^a-zA-Z0-9\s]', '', regex=True)
        
        return result


class AIPredictiveAnalyzer:
    """AI-powered predictive analytics"""
    
    @staticmethod
    def predict_missing_values(df: pd.DataFrame, target_column: str) -> pd.Series:
        """Predict missing values using ML"""
        if not SKLEARN_AVAILABLE:
            return df[target_column]
        
        # Separate rows with and without missing values
        df_complete = df[df[target_column].notna()].copy()
        df_missing = df[df[target_column].isna()].copy()
        
        if len(df_missing) == 0:
            return df[target_column]
        
        # Select numeric features
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        if target_column in numeric_cols:
            numeric_cols.remove(target_column)
        
        if len(numeric_cols) == 0:
            return df[target_column]
        
        # Prepare data
        X_train = df_complete[numeric_cols].fillna(0)
        y_train = df_complete[target_column]
        X_pred = df_missing[numeric_cols].fillna(0)
        
        try:
            # Train model
            model = RandomForestRegressor(n_estimators=100, random_state=42, max_depth=5)
            model.fit(X_train, y_train)
            
            # Predict
            predictions = model.predict(X_pred)
            
            # Fill in predictions
            result = df[target_column].copy()
            result.loc[df_missing.index] = predictions
            return result
        except:
            return df[target_column]
    
    @staticmethod
    def forecast_trend(series: pd.Series, periods: int = 5) -> List[float]:
        """Simple trend forecasting"""
        if not SKLEARN_AVAILABLE:
            return []
        
        clean_series = series.dropna()
        if len(clean_series) < 3:
            return []
        
        X = np.arange(len(clean_series)).reshape(-1, 1)
        y = clean_series.values
        
        try:
            model = LinearRegression()
            model.fit(X, y)
            
            future_X = np.arange(len(clean_series), len(clean_series) + periods).reshape(-1, 1)
            predictions = model.predict(future_X)
            
            return predictions.tolist()
        except:
            return []
    
    @staticmethod
    def cluster_data(df: pd.DataFrame, n_clusters: int = 3) -> np.ndarray:
        """Perform k-means clustering on numeric data"""
        if not SKLEARN_AVAILABLE:
            return np.array([])
        
        numeric_df = df.select_dtypes(include=[np.number]).fillna(0)
        
        if numeric_df.shape[1] == 0:
            return np.array([])
        
        try:
            scaler = StandardScaler()
            scaled_data = scaler.fit_transform(numeric_df)
            
            kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
            clusters = kmeans.fit_predict(scaled_data)
            
            return clusters
        except:
            return np.array([])


class AIFormulaAssistant:
    """AI-powered formula suggestions and generation"""
    
    @staticmethod
    def suggest_formulas(df: pd.DataFrame, column_types: Dict[str, str]) -> List[Dict[str, str]]:
        """Suggest useful formulas based on data structure"""
        suggestions = []
        
        numeric_cols = [col for col, dtype in column_types.items() if dtype in ['numeric', 'currency']]
        
        # Sum suggestions
        if len(numeric_cols) >= 1:
            for col in numeric_cols[:3]:
                suggestions.append({
                    'name': f'Total {col}',
                    'formula': f'SUM([{col}])',
                    'description': f'Calculate total of {col}'
                })
        
        # Average suggestions
        if len(numeric_cols) >= 1:
            for col in numeric_cols[:2]:
                suggestions.append({
                    'name': f'Average {col}',
                    'formula': f'MEAN([{col}])',
                    'description': f'Calculate average of {col}'
                })
        
        # Multiplication (for calculations)
        if len(numeric_cols) >= 2:
            col1, col2 = numeric_cols[0], numeric_cols[1]
            suggestions.append({
                'name': f'{col1} × {col2}',
                'formula': f'[{col1}] * [{col2}]',
                'description': f'Multiply {col1} by {col2}'
            })
        
        # Percentage calculations
        if len(numeric_cols) >= 2:
            col1, col2 = numeric_cols[0], numeric_cols[1]
            suggestions.append({
                'name': f'{col1} as % of {col2}',
                'formula': f'([{col1}] / [{col2}]) * 100',
                'description': f'Calculate {col1} as percentage of {col2}'
            })
        
        # Text concatenation
        text_cols = [col for col, dtype in column_types.items() if dtype in ['text', 'person_name']]
        if len(text_cols) >= 2:
            col1, col2 = text_cols[0], text_cols[1]
            suggestions.append({
                'name': f'Combine {col1} and {col2}',
                'formula': f'CONCAT([{col1}], " ", [{col2}])',
                'description': f'Combine {col1} and {col2} with space'
            })
        
        return suggestions
    
    @staticmethod
    def explain_formula(formula: str) -> str:
        """Explain what a formula does in plain English"""
        explanations = []
        
        # Basic operations
        if '+' in formula:
            explanations.append("adds values together")
        if '-' in formula:
            explanations.append("subtracts values")
        if '*' in formula:
            explanations.append("multiplies values")
        if '/' in formula:
            explanations.append("divides values")
        
        # Functions
        if 'SUM' in formula.upper():
            explanations.append("calculates the sum")
        if 'MEAN' in formula.upper() or 'AVG' in formula.upper():
            explanations.append("calculates the average")
        if 'MAX' in formula.upper():
            explanations.append("finds the maximum value")
        if 'MIN' in formula.upper():
            explanations.append("finds the minimum value")
        if 'IF' in formula.upper():
            explanations.append("applies conditional logic")
        if 'CONCAT' in formula.upper():
            explanations.append("combines text")
        
        if not explanations:
            return "This formula performs a custom calculation"
        
        return "This formula " + ", ".join(explanations) + "."


# Main AI Manager class
class AIManager:
    """Main manager for all AI features"""
    
    def __init__(self, df: pd.DataFrame = None):
        self.df = df
        self.analyzer = None
        self.analysis_cache = None
        
        if df is not None:
            self.set_dataframe(df)
    
    def set_dataframe(self, df: pd.DataFrame):
        """Update the dataframe and reset cache"""
        self.df = df
        self.analyzer = AIDataAnalyzer(df)
        self.analysis_cache = None
    
    def get_full_analysis(self) -> Dict[str, Any]:
        """Get or retrieve cached full analysis"""
        if self.analysis_cache is None and self.analyzer is not None:
            self.analysis_cache = self.analyzer.analyze()
        return self.analysis_cache or {}
    
    def get_column_suggestions(self, column: str) -> List[Dict[str, Any]]:
        """Get AI suggestions for a specific column"""
        if self.df is None or column not in self.df.columns:
            return []
        
        suggestions = []
        series = self.df[column]
        
        # Missing data suggestions
        missing_count = series.isna().sum()
        if missing_count > 0:
            suggestions.append({
                'type': 'cleaning',
                'action': 'fill_missing',
                'description': f'Fill {missing_count} missing values',
                'icon': '🔧'
            })
        
        # Outlier suggestions
        if pd.api.types.is_numeric_dtype(series):
            outliers = AIDataAnalyzer(self.df)._detect_outliers_zscore(series)
            if outliers > 0:
                suggestions.append({
                    'type': 'cleaning',
                    'action': 'remove_outliers',
                    'description': f'Handle {outliers} outliers',
                    'icon': '⚠️'
                })
        
        # Text cleaning suggestions
        if series.dtype == 'object':
            if series.str.strip().ne(series).any():
                suggestions.append({
                    'type': 'cleaning',
                    'action': 'trim_whitespace',
                    'description': 'Remove extra whitespace',
                    'icon': '✂️'
                })
        
        return suggestions
    
    def is_available(self) -> bool:
        """Check if AI features are available"""
        return SKLEARN_AVAILABLE
    
    def get_status_message(self) -> str:
        """Get status message about AI availability"""
        if SKLEARN_AVAILABLE:
            return "✓ AI Features Active"
        else:
            return "⚠ Install scikit-learn for full AI features"


# Export main classes
__all__ = [
    'AIManager',
    'AIDataAnalyzer', 
    'AIDataCleaner',
    'AIPredictiveAnalyzer',
    'AIFormulaAssistant'
]
