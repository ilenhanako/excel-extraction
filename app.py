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
            
            # Find created database files
            db_files = list(files_dir.glob('*.db'))
            self.db_files = db_files
            
            if not db_files:
                return {"error": "No database files created"}
            
            # Extract data from database
            extraction_data = self._extract_from_database(db_files[0])
            
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
        """Extract structured data from SQLite database."""
        try:
            # Query all data
            result = subprocess.run([
                'eparse', '-i', f'sqlite3:///{db_file}', '-o', 'stdout:///', 'query'
            ], capture_output=True, text=True, check=True)
            
            # Parse the output (assuming it's in a structured format)
            data = self._parse_eparse_output(result.stdout)
            
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

def create_sample_excel() -> str:
    """Create a sample Excel file for demonstration."""
    # Employee data
    employee_data = {
        'Name': ['Alice Johnson', 'Bob Smith', 'Charlie Brown', 'Diana Prince', 'Eve Wilson'],
        'Age': [25, 30, 35, 28, 32],
        'City': ['New York', 'Los Angeles', 'Chicago', 'Boston', 'Seattle'],
        'Salary': [75000, 85000, 90000, 80000, 95000],
        'Department': ['Engineering', 'Marketing', 'Sales', 'HR', 'Engineering']
    }
    
    # Sales data
    sales_data = {
        'Product': ['Laptop', 'Phone', 'Tablet', 'Monitor', 'Keyboard'],
        'Q1_Sales': [120, 200, 85, 45, 60],
        'Q2_Sales': [135, 180, 95, 50, 65],
        'Q3_Sales': [110, 220, 90, 40, 70],
        'Q4_Sales': [125, 190, 100, 55, 75]
    }
    
    # Financial data
    financial_data = {
        'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        'Revenue': [50000, 55000, 60000, 65000, 70000, 75000],
        'Expenses': [40000, 42000, 45000, 48000, 50000, 52000],
        'Profit': [10000, 13000, 15000, 17000, 20000, 23000]
    }
    
    filename = "sample_data.xlsx"
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        pd.DataFrame(employee_data).to_excel(writer, sheet_name='Employees', index=False)
        pd.DataFrame(sales_data).to_excel(writer, sheet_name='Sales', index=False)
        pd.DataFrame(financial_data).to_excel(writer, sheet_name='Financial', index=False)
    
    return filename

def process_excel_file(file) -> tuple:
    """Process uploaded Excel file and return extracted data."""
    if file is None:
        return "Please upload an Excel file.", None, None, None, None
    
    try:
        # Initialize extractor
        extractor = ExcelExtractor()
        
        # Extract data
        result = extractor.extract_from_excel(file.name)
        
        if "error" in result:
            return f"Error: {result['error']}", None, None, None, None
        
        # Create visualizations
        charts = create_visualizations(result["data"])
        
        # Format data for display
        summary_text = format_summary(result["data"])
        
        return summary_text, charts[0], charts[1], charts[2], charts[3]
        
    except Exception as e:
        return f"Error processing file: {str(e)}", None, None, None, None

def create_visualizations(data: Dict[str, Any]) -> List[go.Figure]:
    """Create various visualizations from extracted data."""
    charts = []
    
    try:
        # Chart 1: Data Overview
        fig1 = go.Figure()
        fig1.add_trace(go.Indicator(
            mode="gauge+number+delta",
            value=data.get("total_rows", 0),
            title={'text': "Total Data Points"},
            gauge={'axis': {'range': [None, max(data.get("total_rows", 0) * 1.2, 1)]}}
        ))
        fig1.update_layout(title="Data Overview", height=300)
        charts.append(fig1)
        
        # Chart 2: Sheets Distribution (if available)
        if data.get("sheets"):
            fig2 = px.bar(
                x=list(data["sheets"]),
                y=[len(data["sheets"])] * len(data["sheets"]),
                title="Sheets Found",
                labels={'x': 'Sheet Name', 'y': 'Count'}
            )
            fig2.update_layout(height=300)
        else:
            fig2 = go.Figure()
            fig2.add_annotation(text="No sheet data available", xref="paper", yref="paper", x=0.5, y=0.5)
            fig2.update_layout(title="Sheets Found", height=300)
        charts.append(fig2)
        
        # Chart 3: Columns Distribution
        if data.get("columns"):
            fig3 = px.bar(
                x=list(data["columns"]),
                y=[len(data["columns"])] * len(data["columns"]),
                title="Columns Found",
                labels={'x': 'Column Name', 'y': 'Count'}
            )
            fig3.update_layout(height=300)
        else:
            fig3 = go.Figure()
            fig3.add_annotation(text="No column data available", xref="paper", yref="paper", x=0.5, y=0.5)
            fig3.update_layout(title="Columns Found", height=300)
        charts.append(fig3)
        
        # Chart 4: Sample Data Table
        fig4 = go.Figure(data=[go.Table(
            header=dict(values=["Metric", "Value"]),
            cells=dict(values=[
                ["Total Rows", "Sheets", "Columns"],
                [data.get("total_rows", 0), len(data.get("sheets", [])), len(data.get("columns", []))]
            ])
        )])
        fig4.update_layout(title="Data Summary", height=300)
        charts.append(fig4)
        
    except Exception as e:
        # Create empty charts if visualization fails
        for i in range(4):
            fig = go.Figure()
            fig.add_annotation(text=f"Chart {i+1}: Error creating visualization", xref="paper", yref="paper", x=0.5, y=0.5)
            fig.update_layout(title=f"Chart {i+1}", height=300)
            charts.append(fig)
    
    return charts

def format_summary(data: Dict[str, Any]) -> str:
    """Format extracted data into a readable summary."""
    summary = f"""
## Excel Data Extraction Summary

### 📊 Data Overview
- **Total Data Points**: {data.get('total_rows', 0)}
- **Number of Sheets**: {len(data.get('sheets', []))}
- **Number of Columns**: {len(data.get('columns', []))}

### 📋 Sheets Found
{chr(10).join([f"- {sheet}" for sheet in data.get('sheets', [])])}

### 🏷️ Columns Identified
{chr(10).join([f"- {col}" for col in data.get('columns', [])])}

### 🔍 Extraction Details
The data has been successfully extracted and chunked for analysis. 
Each data point maintains its relationship with headers and sheet information.

### 💡 Next Steps
- Use the visualizations below to explore the data structure
- Query specific data points using the extracted information
- Export data for further analysis in other tools
"""
    return summary

def download_sample():
    """Create and return a sample Excel file for download."""
    filename = create_sample_excel()
    return filename

# Create Gradio interface
def create_interface():
    """Create the Gradio interface."""
    
    with gr.Blocks(title="Excel Data Extraction & Visualization", theme=gr.themes.Soft()) as demo:
        gr.Markdown("# 📊 Excel Data Extraction & Visualization Tool")
        gr.Markdown("Upload an Excel file to extract and visualize its data using eparse")
        
        with gr.Row():
            with gr.Column(scale=1):
                file_input = gr.File(
                    label="Upload Excel File",
                    file_types=[".xlsx", ".xls"],
                    type="file"
                )
                
                with gr.Row():
                    upload_btn = gr.Button("📤 Process File", variant="primary")
                    sample_btn = gr.Button("📥 Download Sample", variant="secondary")
                
                gr.Markdown("### Instructions")
                gr.Markdown("""
                1. Upload an Excel file (.xlsx or .xls)
                2. Click 'Process File' to extract data
                3. View the extracted data and visualizations below
                4. Use 'Download Sample' to get a test file
                """)
            
            with gr.Column(scale=2):
                summary_output = gr.Markdown(label="Extraction Summary")
        
        with gr.Row():
            chart1 = gr.Plot(label="Data Overview", show_label=True)
            chart2 = gr.Plot(label="Sheets Distribution", show_label=True)
        
        with gr.Row():
            chart3 = gr.Plot(label="Columns Distribution", show_label=True)
            chart4 = gr.Plot(label="Data Summary Table", show_label=True)
        
        # Event handlers
        upload_btn.click(
            fn=process_excel_file,
            inputs=[file_input],
            outputs=[summary_output, chart1, chart2, chart3, chart4]
        )
        
        sample_btn.click(
            fn=download_sample,
            outputs=[gr.File(label="Sample Excel File")]
        )
        
        gr.Markdown("---")
        gr.Markdown("### About This Tool")
        gr.Markdown("""
        This tool uses **eparse** to extract structured data from Excel files and provides:
        - **Data Extraction**: Automatic table detection and data parsing
        - **Visualization**: Interactive charts showing data structure
        - **Chunking**: Intelligent data segmentation for analysis
        - **Export**: Easy data export for further processing
        
        Built with Gradio and eparse for seamless Excel data exploration.
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