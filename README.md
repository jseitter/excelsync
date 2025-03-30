# ExcelSync

A Python library for managing Excel sheets with predefined structures and validations.

## Features

- Define and validate Excel sheet layouts
- Check Excel sheet structure integrity
- Extract sheet structure to external file format
- Compare Excel sheet structure against stored representation
- Convert Excel content to YAML with schema information
- Support for custom header row positions

## Installation

```bash
pip install excelsync
```

## Usage

```python
from excelsync import ExcelSync

# Load an Excel file
excel = ExcelSync("your_excel_file.xlsx")

# Validate the structure
is_valid, issues = excel.validate_structure()

# Export structure to file
excel.export_structure("structure.json")

# Compare with stored structure
is_matching = excel.compare_structure("structure.json")

# Export to YAML with schema
excel.export_to_yaml("output.yaml")
```

### Working with Custom Header Rows

If your Excel file has headers in a row other than the first row, you can specify the header row:

```python
# Load an Excel file with headers in row 3
excel = ExcelSync("your_excel_file.xlsx", header_row=3)

# Or override the header row for a specific operation
structure = excel.extract_structure(header_row=3)
excel.export_to_yaml("output.yaml", header_row=3)

# The header row information is preserved in the structure
excel.export_structure("structure.json")

# When loading a structure, the header_row is updated automatically
new_excel = ExcelSync("another_file.xlsx")
new_excel.load_structure("structure.json")  # Now new_excel.header_row will be 3
```

## Development

Setup your development environment:

```bash
# Clone the repository
git clone https://github.com/jseitter/excelsync.git
cd excelsync

# Create a virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install development dependencies
pip install -e ".[dev]"

# Run tests
pytest
```

## License

MIT 