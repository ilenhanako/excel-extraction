#!/usr/bin/env python3
"""
Gradio Web App for Excel Data Extraction and Visualization
"""

import gradio as gr
import pandas as pd
import subprocess
import os
import json
import tempfile
from pathlib import Path
from typing import List, Dict, Any
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import re

def extract_document_text(output: str) -> str:
    lines = output.strip().split('\n')
    doc_lines = []
    for line in lines:
        if 'table ' in line and 'found at' in line:
            doc_lines.append(line)
    # If no table lines found, show all output for debugging
    return '\n'.join(doc_lines) if doc_lines else output.strip()

class ExcelExtractor:
    """Class to handle Excel file extraction using eparse."""
    
    def __init__(self):
        self.temp_dir = Path(tempfile.mkdtemp())
        self.db_files = []
    
    def extract_from_excel(self, file_path: str) -> Dict[str, Any]:
        """Extract data from Excel file using eparse."""
        try:
            # Create .files directory for database storage
            files_dir = self.temp_dir / ".files"
            files_dir.mkdir(parents=True, exist_ok=True)
            
            # Parse to SQLite database
            db_file = files_dir / "mydb.db"
            eparse_output = f"sqlite3:///{db_file}"
            print("files_dir:", files_dir)
            print("eparse_output:", eparse_output)
            print("files_dir exists?", files_dir.exists())
            print("files_dir is dir?", files_dir.is_dir())
            print("files_dir writable?", os.access(files_dir, os.W_OK))
            result = subprocess.run([
                'eparse', '-f', file_path, '-o', eparse_output, 'parse', '-z'
            ], capture_output=True, text=True, check=True)
            
            print("eparse stdout:", result.stdout)
            print("eparse stderr:", result.stderr)
            print("eparse returncode:", result.returncode)
            
            print("RAW EPARSE OUTPUT:")
            print(result.stdout)
            
            # Find created database files
            db_files = list(files_dir.glob('*.db'))
            self.db_files = db_files
            
            if not db_files:
                return {"error": "No database files created"}
            
            # Extract data from database
            extraction_data = self._extract_from_database(db_files[0])
            # Add document text extraction from eparse output
            extraction_data["document_text"] = extract_document_text(result.stdout)
            
            return {
                "success": True,
                "data": extraction_data,
                "db_file": str(db_files[0]),
                "message": "Data extracted successfully"
            }
            
        except subprocess.CalledProcessError as e:
            return {"error": f"eparse error: {e.stderr}"}
        except Exception as e:
            return {"error": f"General error: {str(e)}"}
    
    def _extract_from_database(self, db_file: Path) -> Dict[str, Any]:
        """Extract structured data from SQLite database and return both structured and plain text data."""
        try:
            # Query all data
            result = subprocess.run([
                'eparse', '-i', f'sqlite3:///{db_file}', '-o', 'stdout:///', 'query'
            ], capture_output=True, text=True, check=True)
            # Parse the output (assuming it's in a structured format)
            data = self._parse_eparse_output(result.stdout)
            # Add plain text version of all data for display
            data["plain_text"] = result.stdout.strip()
            return data
        except subprocess.CalledProcessError as e:
            return {"error": f"Database query error: {e.stderr}"}
    
    def _parse_eparse_output(self, output: str) -> Dict[str, Any]:
        """Parse eparse output into structured data."""
        lines = output.strip().split('\n')
        
        # Extract headers and data
        data = {
            "tables": [],
            "sheets": set(),
            "columns": set(),
            "total_rows": 0
        }
        
        # Simple parsing - in a real app you'd want more sophisticated parsing
        for line in lines:
            if line.strip() and not line.startswith('==='):
                data["total_rows"] += 1
                # Extract sheet and column info if available
                if 'sheet:' in line:
                    sheet_name = line.split('sheet:')[-1].split()[0]
                    data["sheets"].add(sheet_name)
                if 'c_header:' in line:
                    col_name = line.split('c_header:')[-1].split()[0]
                    data["columns"].add(col_name)
        
        data["sheets"] = list(data["sheets"])
        data["columns"] = list(data["columns"])
        
        return data

def process_excel_file(file) -> str:
    """Process uploaded Excel file and return all extracted data as plain text."""
    if file is None:
        return "Please upload an Excel file."
    try:
        extractor = ExcelExtractor()
        result = extractor.extract_from_excel(file.name)
        if "error" in result:
            return f"Error: {result['error']}"
        # Show all extracted data as plain text
        plain_text = result["data"].get("plain_text", "No data extracted.")
        return plain_text
    except Exception as e:
        return f"Error processing file: {str(e)}"

# Create Gradio interface
def create_interface():
    """Create the Gradio interface (document text only, no sample file)."""
    with gr.Blocks(title="Excel Data Extraction", theme=gr.themes.Soft()) as demo:
        gr.Markdown("# 📄 Excel Document Text Extraction Tool")
        gr.Markdown("Upload an Excel file to extract document text using eparse")
        with gr.Row():
            with gr.Column(scale=1):
                file_input = gr.File(
                    label="Upload Excel File",
                    file_types=[".xlsx", ".xls"],
                    type="file"
                )
                with gr.Row():
                    upload_btn = gr.Button("📤 Process File", variant="primary")
                gr.Markdown("### Instructions")
                gr.Markdown("""
                1. Upload an Excel file (.xlsx or .xls)
                2. Click 'Process File' to extract document text
                3. View the extracted document text below
                """)
            with gr.Column(scale=2):
                document_textbox = gr.Textbox(label="Extracted Document Text", lines=20, interactive=False)
        upload_btn.click(
            fn=process_excel_file,
            inputs=[file_input],
            outputs=[document_textbox]
        )
        gr.Markdown("---")
        gr.Markdown("### About This Tool")
        gr.Markdown("""
        This tool uses **eparse** to extract document text from Excel files.
        Built with Gradio and eparse for seamless Excel document exploration.
        """)
    return demo

if __name__ == "__main__":
    #### Running on local URL:  http://127.0.0.1:7860   OR   http://localhost:7860

    # Create and launch the interface
    demo = create_interface()
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=True,
        show_error=True
    ) 