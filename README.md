# Translation Converter

This tool converts translations from Excel files to JSON format.

## Requirements
- Python 3.8 or higher
- Dependencies listed in requirements.txt

## Installation
1. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage
Run the script with:
```bash
python translation_converter.py
```

The script will:
1. Read the Excel file containing translations
2. Process the translations
3. Generate a new JSON file with the translations 