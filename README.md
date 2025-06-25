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

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Command Line Usage

#### Basic Excel Parsing
```bash
# Parse Excel files and output to console
eparse -f <path_to_excel_files> -o stdout:/// parse

# Parse with verbose output to see all tables found
eparse -f <path_to_excel_files> -v parse
```

#### Store Data in Database
```bash
# Create a SQLite database
mkdir .files
eparse -f <path_to_excel_files> -o sqlite3:/// parse -z
```

#### Query Extracted Data
```bash
# Query all data from database
eparse -i sqlite3:///.files/<db_file> -o stdout:/// query

# Filter data by filename
eparse -i sqlite3:///.files/<db_file> -o stdout:/// query --filter f_name "myfile.xlsx"
```

### Web Application (Gradio)

This project includes two Gradio web applications for interactive Excel data extraction and visualization:

#### 🚀 Quick Start
```bash
# Run the launcher to choose between apps
python run_app.py

# Or run directly:
python app.py              # Basic app
python app_enhanced.py     # Enhanced app
```

#### 📊 Basic App Features
- **File Upload**: Upload Excel files (.xlsx, .xls)
- **Data Extraction**: Extract tables using eparse
- **Basic Visualizations**: Simple charts and summaries
- **Sample Data**: Download test files

#### 🚀 Enhanced App Features
- **Advanced Extraction**: Full table detection with detailed parsing
- **Interactive Visualizations**: Multiple chart types (gauge, bar, pie, tables)
- **Data Type Analysis**: Type inference and distribution
- **Sheet-wise Analysis**: Per-sheet data breakdown
- **Sample Data Preview**: View actual extracted data
- **Comprehensive Summary**: Detailed extraction reports

#### 🎯 Web App Capabilities
- **Upload Interface**: Drag-and-drop Excel file upload
- **Real-time Processing**: Instant data extraction and visualization
- **Interactive Charts**: Plotly-based interactive visualizations
- **Data Export**: Download extracted data and sample files
- **Error Handling**: Graceful error handling and user feedback

### Integration with Unstructured Library
```python
from eparse.contrib.unstructured.partition import partition

# Extract elements from Excel files
elements = partition(filename='your_file.xlsx', eparse_mode='...')
```

## Key Features

- **Table Detection**: Automatically finds tables in Excel sheets
- **Header Preservation**: Maintains relationships between headers and data
- **Multiple Output Formats**: Console, SQLite, PostgreSQL, Web UI
- **Query Capabilities**: Filter and search extracted data
- **Integration**: Works with the `unstructured` library for advanced processing
- **Web Interface**: Interactive Gradio applications for easy data exploration

## Project Structure

```
excel-extraction/
├── README.md
├── requirements.txt
├── app.py                    # Basic Gradio web app
├── app_enhanced.py           # Enhanced Gradio web app
├── run_app.py               # App launcher script
├── examples/
│   └── extraction_example.py # Command line example
└── docs/
    └── research_notes.md     # Detailed research notes
```

## Web App Screenshots

### Basic App
- Simple file upload interface
- Basic data extraction and visualization
- Sample file download

### Enhanced App
- Advanced data extraction with detailed parsing
- Multiple interactive visualizations
- Comprehensive data analysis
- Sample data preview tables

## Contributing

This project is focused on researching and implementing Excel data extraction techniques using eparse and LLM integration.
