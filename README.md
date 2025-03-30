# ExcelSync

A Python library for managing Excel sheets with predefined structures and validations.

## Features

- Define and validate Excel sheet layouts
- Check Excel sheet structure integrity
- Extract sheet structure to external file format
- Compare Excel sheet structure against stored representation
- Convert Excel content to YAML with schema information

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

## Development

Setup your development environment:

```bash
# Clone the repository
git clone https://github.com/yourusername/excelsync.git
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