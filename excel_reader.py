"""
ExcelReader class for efficient Excel file reading using pandas DataFrame.
Loads Excel file once into DataFrame and provides methods to read cells efficiently.
"""

import re
from typing import Any, Optional, List, Dict
import pandas as pd
from utils import get_local_path, is_blank_or_na, cleanup_temp_file


def _parse_cell_address(cell_address: str) -> tuple[int, int]:
    """
    Parse Excel cell address (e.g., 'A1', 'B5') to (row, col) indices.
    
    Args:
        cell_address: Excel cell address like 'A1', 'B5', 'AA10'
    
    Returns:
        Tuple of (row_index, column_index) where both are 0-based
    
    Raises:
        ValueError: If cell address format is invalid
    """
    match = re.match(r'^([A-Z]+)(\d+)$', cell_address.upper())
    if not match:
        raise ValueError(f"Invalid cell address format: {cell_address}")
    
    col_str = match.group(1)
    row_str = match.group(2)
    
    # Convert column letters to number (A=0, B=1, ..., Z=25, AA=26, etc.)
    col_num = 0
    for char in col_str:
        col_num = col_num * 26 + (ord(char) - ord('A') + 1)
    col_num -= 1  # Make it 0-based
    
    # Convert row string to number (1-based to 0-based)
    row_num = int(row_str) - 1
    
    return row_num, col_num


class ExcelReader:
    """
    Efficient Excel file reader that loads file once into DataFrame.
    Supports both local files and S3 URIs.
    """
    
    def __init__(self, excel_path: str, sheet_name: Optional[str] = None):
        """
        Initialize ExcelReader and load Excel file into DataFrame.
        
        Args:
            excel_path: Path to Excel file (local path or S3 URI)
            sheet_name: Optional sheet name. If None, loads first sheet.
        
        Raises:
            FileNotFoundError: If file not found
            ValueError: If sheet name not found
        """
        self.excel_path = excel_path
        self.local_path = None
        self.is_temp = False
        self.sheet_name = sheet_name
        self.dataframe: Optional[pd.DataFrame] = None
        self._load_dataframe()
    
    def _load_dataframe(self) -> None:
        """Load Excel file into pandas DataFrame."""
        self.local_path, self.is_temp = get_local_path(self.excel_path)
        
        try:
            # Read all sheets to get sheet names
            excel_file = pd.ExcelFile(self.local_path)
            
            if self.sheet_name:
                if self.sheet_name not in excel_file.sheet_names:
                    available = excel_file.sheet_names
                    raise ValueError(
                        f"Sheet '{self.sheet_name}' not found. "
                        f"Available sheets: {available}"
                    )
                sheet_to_load = self.sheet_name
            else:
                sheet_to_load = excel_file.sheet_names[0]
                self.sheet_name = sheet_to_load
            
            # Load the sheet into DataFrame
            self.dataframe = pd.read_excel(
                self.local_path,
                sheet_name=sheet_to_load,
                header=None,
                engine='openpyxl'
            )
        except Exception as e:
            if self.is_temp:
                cleanup_temp_file(self.local_path)
            raise
    
    def read_cell(self, cell_address: str) -> Any:
        """
        Read a specific cell from the loaded DataFrame.
        
        Args:
            cell_address: Cell address (e.g., 'A1', 'B5', 'C10')
        
        Returns:
            The cell value, or None if cell is empty
        
        Raises:
            ValueError: If cell address is invalid or out of bounds
        
        Example:
            reader = ExcelReader('sample_data.xlsx')
            value = reader.read_cell('C2')
        """
        if self.dataframe is None:
            raise RuntimeError("DataFrame not loaded")
        
        row_idx, col_idx = _parse_cell_address(cell_address)
        
        if row_idx < 0 or col_idx < 0:
            raise ValueError(f"Invalid cell address: {cell_address}")
        
        if row_idx >= len(self.dataframe) or col_idx >= len(self.dataframe.columns):
            return None
        
        value = self.dataframe.iloc[row_idx, col_idx]
        
        # Convert pandas NaN to None
        if pd.isna(value):
            return None
        
        return value
    
    def read_cell_range(self, start_cell: str, end_cell: str) -> List[Any]:
        """
        Read a range of cells from the loaded DataFrame.
        
        Args:
            start_cell: Starting cell address (e.g., 'A1')
            end_cell: Ending cell address (e.g., 'C10')
        
        Returns:
            List of cell values in row-major order
        
        Example:
            reader = ExcelReader('sample_data.xlsx')
            values = reader.read_cell_range('B2', 'C5')
        """
        if self.dataframe is None:
            raise RuntimeError("DataFrame not loaded")
        
        start_row, start_col = _parse_cell_address(start_cell)
        end_row, end_col = _parse_cell_address(end_cell)
        
        # Ensure valid range
        if start_row > end_row or start_col > end_col:
            raise ValueError(
                f"Invalid range: {start_cell} to {end_cell}. "
                f"Start must be before end."
            )
        
        values = []
        for row_idx in range(start_row, end_row + 1):
            for col_idx in range(start_col, end_col + 1):
                if (row_idx < len(self.dataframe) and 
                    col_idx < len(self.dataframe.columns)):
                    value = self.dataframe.iloc[row_idx, col_idx]
                    if pd.isna(value):
                        values.append(None)
                    else:
                        values.append(value)
                else:
                    values.append(None)
        
        return values
    
    def is_cell_blank(self, cell_address: str) -> bool:
        """
        Check if a specific cell is blank, null, or N/A.
        Returns only True or False.
        
        Args:
            cell_address: Cell address (e.g., 'A1', 'B5')
        
        Returns:
            True if cell is blank/null/N/A, False otherwise
        
        Example:
            reader = ExcelReader('sample_data.xlsx')
            if reader.is_cell_blank('C5'):
                # Handle blank cell
        """
        value = self.read_cell(cell_address)
        return is_blank_or_na(value)
    
    def check_cell_value(self, cell_address: str) -> Dict[str, Any]:
        """
        Read a cell and check if it's blank, null, or N/A.
        
        Args:
            cell_address: Cell address (e.g., 'A1', 'B5')
        
        Returns:
            Dictionary with:
            - 'value': The cell value
            - 'is_blank': True if blank/null/N/A
            - 'cell_address': The cell address
            - 'data_type': Type of the value
        
        Example:
            reader = ExcelReader('sample_data.xlsx')
            result = reader.check_cell_value('C5')
        """
        value = self.read_cell(cell_address)
        is_blank = is_blank_or_na(value)
        
        return {
            'value': value,
            'is_blank': is_blank,
            'cell_address': cell_address,
            'data_type': type(value).__name__ if value is not None else 'NoneType'
        }
    
    def read_all_cells_in_column(self, column: str,
                                 start_row: int = 1,
                                 end_row: Optional[int] = None) -> List[Dict[str, Any]]:
        """
        Read all cells in a specific column.
        
        Args:
            column: Column letter (e.g., 'B', 'C')
            start_row: Starting row number (default: 1, Excel 1-based)
            end_row: Ending row number. If None, reads until first empty cell.
        
        Returns:
            List of dictionaries with cell info:
            [{'address': 'B1', 'value': ..., 'is_blank': ..., 'row': ...}, ...]
        
        Example:
            reader = ExcelReader('sample_data.xlsx')
            cells = reader.read_all_cells_in_column('C', start_row=2)
        """
        if self.dataframe is None:
            raise RuntimeError("DataFrame not loaded")
        
        col_idx = _parse_cell_address(f"{column}1")[1]
        max_rows = len(self.dataframe)
        
        results = []
        row = start_row - 1  # Convert to 0-based
        
        while True:
            if end_row and row >= end_row:
                break
            
            if row >= max_rows:
                break
            
            cell_address = f"{column}{row + 1}"  # Convert back to 1-based for address
            
            if row < len(self.dataframe) and col_idx < len(self.dataframe.columns):
                value = self.dataframe.iloc[row, col_idx]
                if pd.isna(value):
                    value = None
            else:
                value = None
            
            # Stop if we hit empty cell and we're past start_row
            if value is None and row >= start_row - 1:
                if row > start_row - 1:
                    break
            
            results.append({
                'address': cell_address,
                'value': value,
                'is_blank': is_blank_or_na(value),
                'row': row + 1  # Return 1-based row number
            })
            
            row += 1
            if row > 10000:  # Safety limit
                break
        
        return results
    
    def read_all_cells_in_row(self, row: int,
                              start_column: str = 'A',
                              end_column: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Read all cells in a specific row.
        
        Args:
            row: Row number (1-based, Excel notation)
            start_column: Starting column letter (default: 'A')
            end_column: Ending column letter. If None, reads until first empty cell.
        
        Returns:
            List of dictionaries with cell info:
            [{'address': 'A2', 'value': ..., 'is_blank': ..., 'column': ...}, ...]
        
        Example:
            reader = ExcelReader('sample_data.xlsx')
            # Read row 2 from column B to end (until empty)
            cells = reader.read_all_cells_in_row(2, start_column='B')
            
            # Read row 5 from column A to column D
            cells = reader.read_all_cells_in_row(5, start_column='A', end_column='D')
        """
        if self.dataframe is None:
            raise RuntimeError("DataFrame not loaded")
        
        row_idx = row - 1  # Convert to 0-based
        start_col_idx = _parse_cell_address(f"{start_column}1")[1]
        max_cols = len(self.dataframe.columns)
        
        if end_column:
            end_col_idx = _parse_cell_address(f"{end_column}1")[1]
        else:
            end_col_idx = None
        
        if row_idx < 0 or row_idx >= len(self.dataframe):
            return []
        
        results = []
        col_idx = start_col_idx
        
        while True:
            if end_col_idx is not None and col_idx > end_col_idx:
                break
            
            if col_idx >= max_cols:
                break
            
            # Convert column index to letter
            col_letter = self._column_index_to_letter(col_idx)
            cell_address = f"{col_letter}{row}"
            
            if col_idx < len(self.dataframe.columns):
                value = self.dataframe.iloc[row_idx, col_idx]
                if pd.isna(value):
                    value = None
            else:
                value = None
            
            # Stop if we hit empty cell and we're past start_column
            if value is None and col_idx > start_col_idx:
                if end_col_idx is None:  # Only stop early if no end_column specified
                    break
            
            results.append({
                'address': cell_address,
                'value': value,
                'is_blank': is_blank_or_na(value),
                'column': col_letter
            })
            
            col_idx += 1
            if col_idx > 10000:  # Safety limit
                break
        
        return results
    
    @staticmethod
    def _column_index_to_letter(col_idx: int) -> str:
        """
        Convert 0-based column index to Excel column letter(s).
        
        Args:
            col_idx: 0-based column index (0 = A, 1 = B, ..., 25 = Z, 26 = AA, etc.)
        
        Returns:
            Excel column letter(s) (e.g., 'A', 'B', 'AA', 'AB')
        """
        result = ""
        col_idx += 1  # Convert to 1-based for calculation
        
        while col_idx > 0:
            col_idx -= 1
            result = chr(ord('A') + (col_idx % 26)) + result
            col_idx //= 26
        
        return result
    
    def get_dataframe(self) -> pd.DataFrame:
        """
        Get the underlying pandas DataFrame.
        
        Returns:
            The pandas DataFrame containing the Excel data
        """
        if self.dataframe is None:
            raise RuntimeError("DataFrame not loaded")
        return self.dataframe
    
    def close(self) -> None:
        """Clean up resources (e.g., temporary files from S3)."""
        if self.is_temp and self.local_path:
            cleanup_temp_file(self.local_path)
            self.local_path = None
            self.is_temp = False
    
    def __enter__(self):
        """Context manager entry."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - cleanup resources."""
        self.close()
    
    def __del__(self):
        """Destructor - cleanup resources."""
        self.close()

