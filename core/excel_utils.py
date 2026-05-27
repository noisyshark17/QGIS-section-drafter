"""
Excel utilities for the Integrated Geospatial Tools plugin
"""

import os
import pandas as pd


# Define Excel engine preference
try:
    import xlsxwriter
    EXCEL_ENGINE = 'xlsxwriter'
except ImportError:
    EXCEL_ENGINE = 'csv'


class SafeExcelWriter:
    """Excel writer with proper resource management"""
    
    def __init__(self, filepath):
        self.filepath = filepath
        self.sheets = {}
        self._workbook = None
        
    def __enter__(self):
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        
    def add_sheet(self, sheet_name, data):
        """Add a sheet with data"""
        # Sanitize sheet name
        safe_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_'))[:31]
        self.sheets[safe_name] = data
        
    def save(self):
        """Save the Excel file using Pandas for efficiency"""
        try:
            if EXCEL_ENGINE == 'xlsxwriter':
                with pd.ExcelWriter(self.filepath, engine='xlsxwriter') as writer:
                    for sheet_name, data in self.sheets.items():
                        # Accept either dict with 'headers'/'rows' or direct DataFrame/list of dicts
                        if isinstance(data, dict) and 'headers' in data and 'rows' in data:
                            df = pd.DataFrame(data['rows'], columns=data['headers'])
                        else:
                            df = pd.DataFrame(data)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                base_dir = os.path.splitext(self.filepath)[0] + "_profiles"
                os.makedirs(base_dir, exist_ok=True)
                for sheet_name, data in self.sheets.items():
                    csv_path = os.path.join(base_dir, f"{sheet_name}.csv")
                    if isinstance(data, dict) and 'headers' in data and 'rows' in data:
                        df = pd.DataFrame(data['rows'], columns=data['headers'])
                    else:
                        df = pd.DataFrame(data)
                    df.to_csv(csv_path, index=False)
        except Exception as e:
            raise Exception(f"Failed to save file: {str(e)}")
            
    def close(self):
        """Close workbook and free resources"""
        if self._workbook:
            try:
                self._workbook.close()
            except:
                pass
            finally:
                self._workbook = None
