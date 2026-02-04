# Excel Reader Library

A Python library for reading Excel files cell-wise with support for local files and AWS S3 URIs. Includes utilities for checking blank, null, or N/A values.

## Features

- **Efficient DataFrame-based reading**: Load Excel file once into pandas DataFrame for fast repeated access
- **Cell-wise reading**: Read individual cells or cell ranges from Excel files
- **S3 support**: Read Excel files directly from AWS S3 buckets
- **Blank detection**: Check if cells contain blank, null, or N/A values
- **Type safety**: Full type hints for better IDE support
- **Error handling**: Comprehensive error handling with clear messages
- **Context manager support**: Use with `with` statement for automatic cleanup

## Installation

```bash
pip install -r requirements.txt
```

### Requirements

- `openpyxl>=3.1.2` - For reading Excel files
- `pandas>=2.0.0` - For DataFrame operations (used by ExcelReader)
- `boto3>=1.28.0` - For S3 support (optional, only needed for S3 URIs)

## Quick Start

### Recommended: ExcelReader Class (Loads Once)

```python
from excel_reader import ExcelReader

# Load Excel file once (works with local files or S3 URIs)
with ExcelReader('sample_data.xlsx') as reader:
    # Read multiple cells efficiently
    value1 = reader.read_cell('C2')
    value2 = reader.read_cell('C3')
    
    # Check if cell is blank
    if reader.is_cell_blank('C5'):
        # Handle blank cell
        pass
    
    # Access the full DataFrame if needed
    df = reader.get_dataframe()
```

### Alternative: Individual Functions (Opens File Each Time)

```python
from read_excel import read_cell, is_cell_blank

# Read a cell from local file (opens file each time)
value = read_cell('sample_data.xlsx', 'C2')

# Check if cell is blank (opens file again)
if is_cell_blank('sample_data.xlsx', 'C5'):
    # Handle blank cell
    pass
```

**Note**: Use `ExcelReader` class when reading multiple cells from the same file for better performance.

## Usage

### ExcelReader Class (Recommended for Multiple Reads)

The `ExcelReader` class loads the Excel file once into a pandas DataFrame, making it much more efficient when reading multiple cells from the same file.

#### Basic Usage

```python
from excel_reader import ExcelReader

# Load Excel file once (works with local files or S3 URIs)
reader = ExcelReader('sample_data.xlsx')

# Read multiple cells efficiently (no file reopening)
value1 = reader.read_cell('C2')
value2 = reader.read_cell('C3')
value3 = reader.read_cell('B5')

# Check if cells are blank
if reader.is_cell_blank('C5'):
    # Handle blank cell
    pass

# Clean up resources (especially important for S3 files)
reader.close()
```

#### Using Context Manager (Recommended)

```python
from excel_reader import ExcelReader

# Automatic cleanup with context manager
with ExcelReader('sample_data.xlsx') as reader:
    value = reader.read_cell('C2')
    if reader.is_cell_blank('C5'):
        # Handle blank cell
        pass
    
    # Access full DataFrame if needed
    df = reader.get_dataframe()
    # DataFrame is automatically available for pandas operations
```

#### Reading from S3

```python
from excel_reader import ExcelReader

# Works seamlessly with S3 URIs
with ExcelReader('s3://my-bucket/data/file.xlsx') as reader:
    value = reader.read_cell('C2')
    values = reader.read_cell_range('B2', 'C5')
    
    # All operations work the same way
    if reader.is_cell_blank('C5'):
        pass
```

#### Reading Specific Sheet

```python
from excel_reader import ExcelReader

# Load specific sheet
with ExcelReader('sample_data.xlsx', sheet_name='Sheet2') as reader:
    value = reader.read_cell('C2')
```

#### All ExcelReader Methods

```python
from excel_reader import ExcelReader

with ExcelReader('sample_data.xlsx') as reader:
    # Read single cell
    value = reader.read_cell('C2')
    
    # Read cell range
    values = reader.read_cell_range('B2', 'C5')
    
    # Check if blank
    is_blank = reader.is_cell_blank('C5')
    
    # Get detailed cell info
    result = reader.check_cell_value('C5')
    
    # Read all cells in column
    cells = reader.read_all_cells_in_column('C', start_row=2, end_row=10)
    
    # Read all cells in row
    cells = reader.read_all_cells_in_row(2, start_column='B')  # Until empty
    cells = reader.read_all_cells_in_row(5, start_column='A', end_column='D')  # Range
    
    # Access full DataFrame
    df = reader.get_dataframe()
```

### Individual Functions (Opens File Each Time)

For single reads or when you don't need to read multiple cells, you can use the individual functions. Note that these open the file each time, so use `ExcelReader` for multiple reads.

### Reading Cells

#### Read a Single Cell

```python
from read_excel import read_cell

# Local file
value = read_cell('sample_data.xlsx', 'C2')
value = read_cell('sample_data.xlsx', 'B5', sheet_name='Sheet1')

# S3 URI
value = read_cell('s3://my-bucket/data/file.xlsx', 'C2')
```

#### Read a Cell Range

```python
from read_excel import read_cell_range

# Returns list of values in row-major order
values = read_cell_range('sample_data.xlsx', 'B2', 'C5')

# S3 URI
values = read_cell_range('s3://my-bucket/data/file.xlsx', 'B2', 'C5')
```

#### Read All Cells in a Column

```python
from read_excel import read_all_cells_in_column

# Read from row 2 to row 12
cells = read_all_cells_in_column('sample_data.xlsx', 'C', start_row=2, end_row=12)

# Read until first empty cell
cells = read_all_cells_in_column('sample_data.xlsx', 'C', start_row=2)

# Returns list of dicts: [{'address': 'C2', 'value': ..., 'is_blank': ..., 'row': 2}, ...]
for cell_info in cells:
    print(f"{cell_info['address']}: {cell_info['value']}")
```

#### Read All Cells in a Row

```python
from read_excel import read_all_cells_in_row

# Read row 2 from column B until first empty cell
cells = read_all_cells_in_row('sample_data.xlsx', 2, start_column='B')

# Read row 5 from column A to column D
cells = read_all_cells_in_row('sample_data.xlsx', 5, start_column='A', end_column='D')

# Read row 3 from column A until empty (default start is 'A')
cells = read_all_cells_in_row('sample_data.xlsx', 3)

# S3 URI
cells = read_all_cells_in_row('s3://my-bucket/data/file.xlsx', 2, start_column='B')

# Returns list of dicts: [{'address': 'B2', 'value': ..., 'is_blank': ..., 'column': 'B'}, ...]
for cell_info in cells:
    print(f"{cell_info['address']}: {cell_info['value']}")
```

#### Loop Through Specific Range (B2 to D2)

```python
from read_excel import read_all_cells_in_row

# Read row 2 from column B to column D
cells = read_all_cells_in_row('sample_data.xlsx', 2, start_column='B', end_column='D')

# Loop through B2, C2, D2
for cell_info in cells:
    column = cell_info['column']  # 'B', 'C', or 'D'
    value = cell_info['value']
    address = cell_info['address']  # 'B2', 'C2', or 'D2'
    is_blank = cell_info['is_blank']
    
    print(f"{address} ({column}): {value}")

# Or using ExcelReader (more efficient for multiple operations)
from excel_reader import ExcelReader

with ExcelReader('sample_data.xlsx') as reader:
    cells = reader.read_all_cells_in_row(2, start_column='B', end_column='D')
    
    for cell_info in cells:
        print(f"Column {cell_info['column']}: {cell_info['value']}")
```

### Checking Blank/Null/N/A Values

#### Check if a Cell is Blank

```python
from read_excel import is_cell_blank

# Returns True or False only
if is_cell_blank('sample_data.xlsx', 'C5'):
    print("Cell is empty or N/A")

# S3 URI
if is_cell_blank('s3://my-bucket/data/file.xlsx', 'C5'):
    print("Cell is empty or N/A")
```

#### Check a Value (Not from Excel)

```python
from utils import is_blank_or_na

# Check any value
if is_blank_or_na(None):
    print("Value is blank")

if is_blank_or_na('N/A'):
    print("Value is N/A")

if is_blank_or_na(''):
    print("Value is empty string")
```

#### Get Detailed Cell Information

```python
from read_excel import check_cell_value

# Returns dictionary with value, is_blank, cell_address, and data_type
result = check_cell_value('sample_data.xlsx', 'C5')

print(result['value'])        # The cell value
print(result['is_blank'])      # True if blank/null/N/A
print(result['cell_address'])  # 'C5'
print(result['data_type'])     # 'str', 'int', 'NoneType', etc.
```

### Utility Functions

#### Check if Path is S3 URI

```python
from utils import is_s3_uri

if is_s3_uri('s3://bucket/file.xlsx'):
    print("This is an S3 URI")
```

#### Download from S3

```python
from utils import download_from_s3

# Download to temporary file (auto-generated)
local_path = download_from_s3('s3://my-bucket/data/file.xlsx')

# Download to specific path
local_path = download_from_s3('s3://my-bucket/data/file.xlsx', '/tmp/file.xlsx')
```

## API Reference

### `excel_reader` Module

#### `ExcelReader(excel_path: str, sheet_name: Optional[str] = None)`

Efficient Excel file reader that loads file once into DataFrame.

**Parameters:**
- `excel_path`: Path to Excel file (local or S3 URI)
- `sheet_name`: Optional sheet name (uses first sheet if None)

**Methods:**

- `read_cell(cell_address: str) -> Any`: Read a specific cell
- `read_cell_range(start_cell: str, end_cell: str) -> List[Any]`: Read a range of cells
- `is_cell_blank(cell_address: str) -> bool`: Check if cell is blank/null/N/A
- `check_cell_value(cell_address: str) -> Dict[str, Any]`: Get detailed cell info
- `read_all_cells_in_column(column: str, start_row: int = 1, end_row: Optional[int] = None) -> List[Dict[str, Any]]`: Read all cells in column
- `read_all_cells_in_row(row: int, start_column: str = 'A', end_column: Optional[str] = None) -> List[Dict[str, Any]]`: Read all cells in row
- `get_dataframe() -> pd.DataFrame`: Get the underlying pandas DataFrame
- `close() -> None`: Clean up resources (called automatically with context manager)

**Context Manager Support:**
```python
with ExcelReader('file.xlsx') as reader:
    # Use reader
    pass
# Automatically cleaned up
```

**Raises:**
- `FileNotFoundError`: If file not found
- `ValueError`: If sheet name not found

### `read_excel` Module

#### `read_cell(excel_path: str, cell_address: str, sheet_name: Optional[str] = None) -> Any`

Read a specific cell from an Excel file.

**Parameters:**
- `excel_path`: Path to Excel file (local or S3 URI)
- `cell_address`: Cell address (e.g., 'A1', 'B5')
- `sheet_name`: Optional sheet name (uses active sheet if None)

**Returns:** Cell value or None

**Raises:**
- `FileNotFoundError`: If file not found
- `ValueError`: If sheet name not found

#### `read_cell_range(excel_path: str, start_cell: str, end_cell: str, sheet_name: Optional[str] = None) -> List[Any]`

Read a range of cells.

**Parameters:**
- `excel_path`: Path to Excel file (local or S3 URI)
- `start_cell`: Starting cell address
- `end_cell`: Ending cell address
- `sheet_name`: Optional sheet name

**Returns:** List of cell values in row-major order

#### `is_cell_blank(excel_path: str, cell_address: str, sheet_name: Optional[str] = None) -> bool`

Check if a cell is blank, null, or N/A. Returns only True or False.

**Parameters:**
- `excel_path`: Path to Excel file (local or S3 URI)
- `cell_address`: Cell address
- `sheet_name`: Optional sheet name

**Returns:** True if blank/null/N/A, False otherwise

#### `check_cell_value(excel_path: str, cell_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]`

Get detailed information about a cell.

**Returns:** Dictionary with keys: `value`, `is_blank`, `cell_address`, `data_type`

#### `read_all_cells_in_column(excel_path: str, column: str, start_row: int = 1, end_row: Optional[int] = None, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]`

Read all cells in a column.

**Parameters:**
- `excel_path`: Path to Excel file (local or S3 URI)
- `column`: Column letter (e.g., 'B', 'C')
- `start_row`: Starting row number (default: 1)
- `end_row`: Ending row number (None = until first empty)
- `sheet_name`: Optional sheet name

**Returns:** List of dictionaries with cell info

#### `read_all_cells_in_row(excel_path: str, row: int, start_column: str = 'A', end_column: Optional[str] = None, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]`

Read all cells in a row.

**Parameters:**
- `excel_path`: Path to Excel file (local or S3 URI)
- `row`: Row number (1-based, Excel notation)
- `start_column`: Starting column letter (default: 'A')
- `end_column`: Ending column letter (None = until first empty)
- `sheet_name`: Optional sheet name

**Returns:** List of dictionaries with cell info: `[{'address': 'A2', 'value': ..., 'is_blank': ..., 'column': 'A'}, ...]`

### `utils` Module

#### `is_blank_or_na(value: Any) -> bool`

Check if a value is blank, null, or N/A.

**Parameters:**
- `value`: Value to check

**Returns:** True if blank/null/N/A, False otherwise

**Recognized blank values:**
- `None`
- Empty string `''`
- `'N/A'`, `'NA'`
- `'NULL'`, `'NONE'`
- `'#N/A'`, `'#NA'`
- Whitespace-only strings

#### `is_s3_uri(path: str) -> bool`

Check if path is an S3 URI.

#### `download_from_s3(s3_uri: str, local_path: Optional[str] = None) -> str`

Download file from S3.

**Returns:** Path to local file

**Raises:**
- `ImportError`: If boto3 not installed
- `RuntimeError`: If AWS credentials not configured
- `FileNotFoundError`: If file or bucket not found

#### `get_local_path(excel_path: str) -> Tuple[str, bool]`

Get local file path, downloading from S3 if necessary.

**Returns:** Tuple of (local_file_path, is_temporary_file)

## AWS S3 Configuration

To use S3 URIs, configure AWS credentials using one of:

1. **AWS CLI:**
   ```bash
   aws configure
   ```

2. **Environment Variables:**
   ```bash
   export AWS_ACCESS_KEY_ID=your_access_key
   export AWS_SECRET_ACCESS_KEY=your_secret_key
   ```

3. **IAM Role:** (if running on EC2/Lambda)

## Examples

### Example 1: Using ExcelReader (Recommended for Multiple Reads)

```python
from excel_reader import ExcelReader

excel_file = 'sample_data.xlsx'

# Load file once, read multiple cells efficiently
with ExcelReader(excel_file) as reader:
    # Read multiple cells (file only opened once)
    name = reader.read_cell('C2')
    email = reader.read_cell('C3')
    age = reader.read_cell('C4')
    
    # Check if optional field is blank
    if not reader.is_cell_blank('C5'):
        phone = reader.read_cell('C5')
    else:
        phone = None
    
    # Read a range
    values = reader.read_cell_range('B2', 'C5')
    
    # Access full DataFrame for advanced operations
    df = reader.get_dataframe()
```

### Example 2: Read and Validate Cells (Individual Functions)

```python
from read_excel import read_cell, is_cell_blank

excel_file = 'sample_data.xlsx'

# Read a cell (opens file)
name = read_cell(excel_file, 'C2')

# Check if optional field is blank (opens file again)
if not is_cell_blank(excel_file, 'C5'):
    phone = read_cell(excel_file, 'C5')
else:
    phone = None
```

### Example 3: Process Column Data

```python
from read_excel import read_all_cells_in_column

cells = read_all_cells_in_column('sample_data.xlsx', 'C', start_row=2)

for cell_info in cells:
    if not cell_info['is_blank']:
        print(f"Row {cell_info['row']}: {cell_info['value']}")
```

### Example 3b: Process Row Data

```python
from read_excel import read_all_cells_in_row

# Read row 2 from column B until empty
cells = read_all_cells_in_row('sample_data.xlsx', 2, start_column='B')

for cell_info in cells:
    if not cell_info['is_blank']:
        print(f"Column {cell_info['column']}: {cell_info['value']}")

# Read specific range
cells = read_all_cells_in_row('sample_data.xlsx', 5, start_column='A', end_column='D')
```

### Example 4: Read from S3 with ExcelReader

```python
from excel_reader import ExcelReader

s3_path = 's3://my-bucket/data/file.xlsx'

# Load from S3 once, read multiple cells
with ExcelReader(s3_path) as reader:
    value = reader.read_cell('C2')
    if reader.is_cell_blank('C5'):
        print("Cell is blank")
    
    # All operations work the same way
    values = reader.read_cell_range('B2', 'C5')
```

### Example 5: Read from S3 (Individual Functions)

```python
from read_excel import read_cell, is_cell_blank

s3_path = 's3://my-bucket/data/file.xlsx'

value = read_cell(s3_path, 'C2')

if is_cell_blank(s3_path, 'C5'):
    print("Cell is blank")
```

### Example 6: Create Sample Excel File

```python
from create_sample_excel import create_sample_excel

# Create sample file
excel_file = create_sample_excel('sample_data.xlsx')
```

## Error Handling

All functions raise appropriate exceptions:

- `FileNotFoundError`: File doesn't exist (local or S3)
- `ValueError`: Invalid sheet name or cell address
- `RuntimeError`: S3 download errors or credential issues
- `ImportError`: Missing dependencies (boto3 for S3)

## Notes

- **Performance**: Use `ExcelReader` class when reading multiple cells from the same file. It loads the file once into a DataFrame, making subsequent reads much faster.
- **Individual Functions**: Use individual functions (`read_cell`, etc.) for single reads or when you don't need to read multiple cells.
- **S3 Files**: S3 files are automatically downloaded to temporary files and cleaned up after use (especially important with `ExcelReader.close()` or context manager).
- **Cell Addresses**: Cell addresses use Excel notation (e.g., 'A1', 'B5', 'AA10').
- **Blank Detection**: Recognizes: None, empty strings, 'N/A', 'NA', 'NULL', 'NONE', '#N/A', '#NA'.
- **DataFrame Access**: With `ExcelReader`, you can access the full pandas DataFrame using `get_dataframe()` for advanced operations.
- **Transparent Support**: All functions and classes support both local paths and S3 URIs transparently.

## License

This project is provided as-is for use in your applications.

