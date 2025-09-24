# Excel MCP Server

A comprehensive Model Context Protocol (MCP) server for Excel and CSV file operations, providing advanced data analysis, formula extraction, and file manipulation capabilities.

## Features

### Data Discovery & Metadata
- **Sheet enumeration**: List all sheet names in Excel files
- **Data structure analysis**: Get detailed information about sheet dimensions, columns, and data types
- **File metadata**: Extract file information including size, modification dates, and sheet counts

### Data Reading & Writing
- **Flexible data reading**: Read specific ranges from Excel sheets and CSV files
- **Data export**: Write data to Excel and CSV formats with customizable options
- **Format conversion**: Convert between Excel and CSV formats seamlessly

### Data Analysis & Quality
- **Statistical analysis**: Generate comprehensive data summaries with descriptive statistics
- **Data quality validation**: Perform thorough quality checks including missing values, duplicates, and outlier detection
- **Search capabilities**: Find specific data within sheets with case-sensitive options

### Formula Analysis
- **Formula extraction**: Identify and extract all formulas from Excel sheets
- **Calculation derivation**: Show how values are calculated with actual examples
- **Pattern recognition**: Analyze formula patterns and identify commonly used functions
- **Reference mapping**: Track cell dependencies and formula relationships

## Installation

1. Clone the repository:
```bash
git clone https://github.com/AltNeo/excel-mcp.git
cd excel-mcp
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the MCP server:
```bash
python excel_mcp.py
```

## Requirements

- Python 3.7+
- pandas
- openpyxl
- fastmcp

## Available Tools

### File Discovery
- `get_excel_sheet_names(file_path)` - List all sheet names in an Excel file
- `get_file_info(file_path)` - Get comprehensive file metadata

### Data Reading
- `read_excel_range(file_path, sheet_name, start_row, end_row, columns)` - Read specific Excel ranges
- `read_csv_range(file_path, start_row, end_row, columns, delimiter)` - Read specific CSV ranges
- `get_sheet_info(file_path, sheet_name)` - Get detailed sheet metadata

### Data Writing & Conversion
- `write_excel_data(file_path, data, sheet_name)` - Write JSON data to Excel
- `write_csv_data(file_path, data, delimiter)` - Write JSON data to CSV
- `convert_excel_to_csv(excel_path, csv_path, sheet_name)` - Convert Excel to CSV
- `convert_csv_to_excel(csv_path, excel_path, sheet_name)` - Convert CSV to Excel

### Data Analysis
- `get_data_summary(file_path, sheet_name, columns)` - Generate statistical summaries
- `search_in_sheet(file_path, search_term, sheet_name, columns)` - Search for data within sheets
- `validate_data_quality(file_path, sheet_name, columns)` - Comprehensive data quality analysis

### Formula Analysis
- `analyze_excel_formulas(file_path, sheet_name, max_examples)` - Extract and analyze Excel formulas with examples

## Usage Examples

### Basic Workflow
```python
# 1. Discover available sheets
sheets = get_excel_sheet_names("data.xlsx")

# 2. Get sheet structure information
info = get_sheet_info("data.xlsx", "Sales")

# 3. Read specific data
data = read_excel_range("data.xlsx", "Sales", start_row=0, end_row=100)

# 4. Analyze data quality
quality_report = validate_data_quality("data.xlsx", "Sales")

# 5. Extract formulas and understand calculations
formula_analysis = analyze_excel_formulas("data.xlsx", "Sales", max_examples=5)
```

### Formula Analysis Output
The formula analysis provides detailed insights including:
- Total number of formulas found
- Columns containing formulas
- Most commonly used Excel functions
- Formula patterns and cell references
- Step-by-step calculation examples showing input values, formulas, and results

## Data Quality Features

The data quality validation includes:
- Missing value analysis with percentages
- Duplicate row detection
- Data type consistency checking
- Outlier identification using IQR method
- Overall quality score (0-100)
- Actionable recommendations for data improvement

## Protocol Integration

This server implements the Model Context Protocol (MCP) and can be integrated with MCP-compatible clients for automated Excel data processing workflows.

## License

MIT License - see LICENSE file for details.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.
