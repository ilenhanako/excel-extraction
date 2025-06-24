# Excel Extraction Project

Research Notes on [Notion](https://www.notion.so/Excel-OCR-21c8bf45dbd08094b2ddc43456dfed2b)

## Key Questions
- How to provide the documents to the LLMs? Performance?
- **How do you get specific data out?**
- Accuracy of the Data?

## Approaches

### Eparse + Unstructured + LLM: Summarizing & Querying Unstructured Data from Excel
[Github](https://github.com/ChrisPappalardo/eparse) to eparse

## Installation

### Prerequisites
- Python 3.9+
- pip

### Setup
1. Create a virtual environment:
```bash
python3 -m venv eparse_env
source eparse_env/bin/activate
```

2. Install eparse:
```bash
pip install eparse
```

## Usage

### Basic Excel Parsing
```bash
# Parse Excel files and output to console
eparse -f <path_to_excel_files> -o stdout:/// parse

# Parse with verbose output to see all tables found
eparse -f <path_to_excel_files> -v parse
```

### Store Data in Database
```bash
# Create a SQLite database
mkdir .files
eparse -f <path_to_excel_files> -o sqlite3:/// parse -z
```

### Query Extracted Data
```bash
# Query all data from database
eparse -i sqlite3:///.files/<db_file> -o stdout:/// query

# Filter data by filename
eparse -i sqlite3:///.files/<db_file> -o stdout:/// query --filter f_name "myfile.xlsx"
```

### Integration with Unstructured Library
```python
from eparse.contrib.unstructured.partition import partition

# Extract elements from Excel files
elements = partition(filename='your_file.xlsx', eparse_mode='...')
```

## Key Features

- **Table Detection**: Automatically finds tables in Excel sheets
- **Header Preservation**: Maintains relationships between row/column headers
- **Multiple Output Formats**: Console, SQLite, PostgreSQL
- **Query Capabilities**: Filter and search extracted data
- **Integration**: Works with the `unstructured` library for advanced processing

## Project Structure

```
excel-extraction/
├── README.md
├── requirements.txt
├── examples/
│   ├── sample_data.xlsx
│   └── extraction_example.py
└── docs/
    └── research_notes.md
```

## Contributing

This project is focused on researching and implementing Excel data extraction techniques using eparse and LLM integration.
