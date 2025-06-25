#!/usr/bin/env python3
"""
Enhanced Gradio: Better handle eparse output and more detailed data extraction
"""

import gradio as gr
import pandas as pd
import subprocess
import os
import json
import tempfile
import sqlite3
from pathlib import Path
from typing import List, Dict, Any, Tuple
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import re

class EnhancedExcelExtractor:
    """Enhanced class to handle Excel file extraction using eparse."""
    
    def __init__(self):
        self.temp_dir = Path(tempfile.mkdtemp())
        self.db_files = []
        self.extracted_data = {}
    
    def extract_from_excel(self, file_path: str):
        """Extract data from Excel file using eparse with enhanced parsing."""
        try:
            # Create .files directory for database storage
            files_dir = self.temp_dir / ".files"
            files_dir.mkdir(exist_ok=True)
            
            # Parse to SQLite database
            result = subprocess.run([
                'eparse', '-f', file_path, '-o', f'sqlite3:///{files_dir}', 'parse', '-z'
            ], capture_output=True, text=True, check=True)
            
            # Find created database files
            db_files = list(files_dir.glob('*.db'))
            self.db_files = db_files
            
            if not db_files:
                return {"error": "No database files created"}
            
            # Extract detailed data from database
            extraction_data = self._extract_detailed_data(db_files[0])
            
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
    
    def _extract_detailed_data(self, db_file: Path):
        """Extract detailed structured data from SQLite database."""
        try:
            # Connect to the database
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            
            # Get table schema
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()
            
            data = {
                "tables": [],
                "sheets": set(),
                "columns": set(),
                "total_rows": 0,
                "data_types": set(),
                "raw_data": [],
                "sheet_data": {},
                "column_data": {}
            }
            
            for table in tables:
                table_name = table[0]
                if table_name == "excelparse":
                    # Get all data from the excelparse table
                    cursor.execute("SELECT * FROM excelparse")
                    rows = cursor.fetchall()
                    
                    # Get column names
                    cursor.execute("PRAGMA table_info(excelparse)")
                    columns = [col[1] for col in cursor.fetchall()]
                    
                    data["total_rows"] = len(rows)
                    
                    # Process each row
                    for row in rows:
                        row_dict = dict(zip(columns, row))
                        data["raw_data"].append(row_dict)
                        
                        # Extract sheet information
                        if row_dict.get('sheet'):
                            data["sheets"].add(row_dict['sheet'])
                            if row_dict['sheet'] not in data["sheet_data"]:
                                data["sheet_data"][row_dict['sheet']] = []
                            data["sheet_data"][row_dict['sheet']].append(row_dict)
                        
                        # Extract column information
                        if row_dict.get('c_header'):
                            data["columns"].add(row_dict['c_header'])
                            if row_dict['c_header'] not in data["column_data"]:
                                data["column_data"][row_dict['c_header']] = []
                            data["column_data"][row_dict['c_header']].append(row_dict)
                        
                        # Extract data types
                        if row_dict.get('type'):
                            data["data_types"].add(row_dict['type'])
            
            conn.close()
            
            # Convert sets to lists for JSON serialization
            data["sheets"] = list(data["sheets"])
            data["columns"] = list(data["columns"])
            data["data_types"] = list(data["data_types"])
            
            return data
            
        except Exception as e:
            return {"error": f"Database extraction error: {str(e)}"}

def create_sample_excel():
    """Create a comprehensive sample Excel file for demonstration."""
    # Employee data
    employee_data = {
        'Name': ['Alice Johnson', 'Bob Smith', 'Charlie Brown', 'Diana Prince', 'Eve Wilson', 'Frank Miller'],
        'Age': [25, 30, 35, 28, 32, 29],
        'City': ['New York', 'Los Angeles', 'Chicago', 'Boston', 'Seattle', 'Austin'],
        'Salary': [75000, 85000, 90000, 80000, 95000, 82000],
        'Department': ['Engineering', 'Marketing', 'Sales', 'HR', 'Engineering', 'Design'],
        'Start_Date': ['2020-01-15', '2019-03-20', '2018-07-10', '2021-02-01', '2020-11-15', '2021-06-01']
    }
    
    # Sales data
    sales_data = {
        'Product': ['Laptop', 'Phone', 'Tablet', 'Monitor', 'Keyboard', 'Mouse'],
        'Q1_Sales': [120, 200, 85, 45, 60, 80],
        'Q2_Sales': [135, 180, 95, 50, 65, 85],
        'Q3_Sales': [110, 220, 90, 40, 70, 90],
        'Q4_Sales': [125, 190, 100, 55, 75, 95],
        'Total_Sales': [490, 790, 370, 190, 270, 350]
    }
    
    # Financial data
    financial_data = {
        'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        'Revenue': [50000, 55000, 60000, 65000, 70000, 75000],
        'Expenses': [40000, 42000, 45000, 48000, 50000, 52000],
        'Profit': [10000, 13000, 15000, 17000, 20000, 23000],
        'Growth_Rate': [0.05, 0.10, 0.09, 0.08, 0.08, 0.07]
    }
    
    filename = "enhanced_sample_data.xlsx"
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        pd.DataFrame(employee_data).to_excel(writer, sheet_name='Employees', index=False)
        pd.DataFrame(sales_data).to_excel(writer, sheet_name='Sales', index=False)
        pd.DataFrame(financial_data).to_excel(writer, sheet_name='Financial', index=False)
    
    return filename

def process_excel_file_enhanced(file):
    """Process uploaded Excel file and return enhanced extracted data."""
    if file is None:
        return "Please upload an Excel file.", None, None, None, None, None, None
    
    try:
        # Initialize enhanced extractor
        extractor = EnhancedExcelExtractor()
        
        # Extract data
        result = extractor.extract_from_excel(file.name)
        
        if "error" in result:
            return f"Error: {result['error']}", None, None, None, None, None, None
        
        # Create enhanced visualizations
        charts = create_enhanced_visualizations(result["data"])
        
        # Format enhanced data for display
        summary_text = format_enhanced_summary(result["data"])
        
        return summary_text, charts[0], charts[1], charts[2], charts[3], charts[4], charts[5]
        
    except Exception as e:
        return f"Error processing file: {str(e)}", None, None, None, None, None, None

def create_enhanced_visualizations(data: Dict[str, Any]):
    """Create enhanced visualizations from extracted data."""
    charts = []
    
    try:
        # Chart 1: Data Overview Gauge
        fig1 = go.Figure()
        fig1.add_trace(go.Indicator(
            mode="gauge+number+delta",
            value=data.get("total_rows", 0),
            title={'text': "Total Data Points"},
            gauge={
                'axis': {'range': [None, max(data.get("total_rows", 0) * 1.2, 1)]},
                'bar': {'color': "darkblue"},
                'steps': [
                    {'range': [0, data.get("total_rows", 0) * 0.3], 'color': "lightgray"},
                    {'range': [data.get("total_rows", 0) * 0.3, data.get("total_rows", 0) * 0.7], 'color': "yellow"},
                    {'range': [data.get("total_rows", 0) * 0.7, data.get("total_rows", 0)], 'color': "green"}
                ]
            }
        ))
        fig1.update_layout(title="Data Overview", height=300)
        charts.append(fig1)
        
        # Chart 2: Sheets Distribution
        if data.get("sheets"):
            sheet_counts = [len(data["sheet_data"].get(sheet, [])) for sheet in data["sheets"]]
            fig2 = px.bar(
                x=data["sheets"],
                y=sheet_counts,
                title="Data Points per Sheet",
                labels={'x': 'Sheet Name', 'y': 'Number of Data Points'},
                color=sheet_counts,
                color_continuous_scale='viridis'
            )
            fig2.update_layout(height=300)
        else:
            fig2 = go.Figure()
            fig2.add_annotation(text="No sheet data available", xref="paper", yref="paper", x=0.5, y=0.5)
            fig2.update_layout(title="Sheets Distribution", height=300)
        charts.append(fig2)
        
        # Chart 3: Columns Distribution
        if data.get("columns"):
            column_counts = [len(data["column_data"].get(col, [])) for col in data["columns"]]
            fig3 = px.bar(
                x=data["columns"],
                y=column_counts,
                title="Data Points per Column",
                labels={'x': 'Column Name', 'y': 'Number of Data Points'},
                color=column_counts,
                color_continuous_scale='plasma'
            )
            fig3.update_layout(height=300)
        else:
            fig3 = go.Figure()
            fig3.add_annotation(text="No column data available", xref="paper", yref="paper", x=0.5, y=0.5)
            fig3.update_layout(title="Columns Distribution", height=300)
        charts.append(fig3)
        
        # Chart 4: Data Types Distribution
        if data.get("data_types"):
            type_counts = {}
            for data_type in data["data_types"]:
                type_counts[data_type] = sum(1 for row in data["raw_data"] if row.get('type') == data_type)
            
            fig4 = px.pie(
                values=list(type_counts.values()),
                names=list(type_counts.keys()),
                title="Data Types Distribution"
            )
            fig4.update_layout(height=300)
        else:
            fig4 = go.Figure()
            fig4.add_annotation(text="No data type information available", xref="paper", yref="paper", x=0.5, y=0.5)
            fig4.update_layout(title="Data Types Distribution", height=300)
        charts.append(fig4)
        
        # Chart 5: Sample Data Table
        if data.get("raw_data"):
            # Create a sample table from the first few rows
            sample_data = data["raw_data"][:10]
            if sample_data:
                headers = list(sample_data[0].keys())
                values = [[row.get(h, '') for h in headers] for row in sample_data]
                
                fig5 = go.Figure(data=[go.Table(
                    header=dict(values=headers, fill_color='paleturquoise', align='left'),
                    cells=dict(values=list(zip(*values)), fill_color='lavender', align='left'))
                ])
                fig5.update_layout(title="Sample Extracted Data (First 10 Rows)", height=400)
            else:
                fig5 = go.Figure()
                fig5.add_annotation(text="No data available", xref="paper", yref="paper", x=0.5, y=0.5)
                fig5.update_layout(title="Sample Data", height=400)
        else:
            fig5 = go.Figure()
            fig5.add_annotation(text="No data available", xref="paper", yref="paper", x=0.5, y=0.5)
            fig5.update_layout(title="Sample Data", height=400)
        charts.append(fig5)
        
        # Chart 6: Data Summary Table
        summary_data = [
            ["Total Rows", data.get("total_rows", 0)],
            ["Sheets", len(data.get("sheets", []))],
            ["Columns", len(data.get("columns", []))],
            ["Data Types", len(data.get("data_types", []))],
            ["Unique Values", len(set(str(row.get('value', '')) for row in data.get("raw_data", [])))]
        ]
        
        fig6 = go.Figure(data=[go.Table(
            header=dict(values=["Metric", "Value"], fill_color='lightblue', align='left'),
            cells=dict(values=list(zip(*summary_data)), fill_color='lightcyan', align='left'))
        ])
        fig6.update_layout(title="Data Summary", height=300)
        charts.append(fig6)
        
    except Exception as e:
        # Create empty charts if visualization fails
        for i in range(6):
            fig = go.Figure()
            fig.add_annotation(text=f"Chart {i+1}: Error creating visualization - {str(e)}", xref="paper", yref="paper", x=0.5, y=0.5)
            fig.update_layout(title=f"Chart {i+1}", height=300)
            charts.append(fig)
    
    return charts

def format_enhanced_summary(data: Dict[str, Any]):
    """Format extracted data into an enhanced readable summary."""
    summary = f"""
## 📊 Enhanced Excel Data Extraction Summary

### 🔢 Data Overview
- **Total Data Points**: {data.get('total_rows', 0):,}
- **Number of Sheets**: {len(data.get('sheets', []))}
- **Number of Columns**: {len(data.get('columns', []))}
- **Data Types Found**: {len(data.get('data_types', []))}
- **Unique Values**: {len(set(str(row.get('value', '')) for row in data.get('raw_data', []))):,}

### 📋 Sheets Found
{chr(10).join([f"- **{sheet}**: {len(data.get('sheet_data', {}).get(sheet, []))} data points" for sheet in data.get('sheets', [])])}

### 🏷️ Columns Identified
{chr(10).join([f"- **{col}**: {len(data.get('column_data', {}).get(col, []))} data points" for col in data.get('columns', [])])}

### 📊 Data Types
{chr(10).join([f"- {dtype}" for dtype in data.get('data_types', [])])}

### 🔍 Extraction Details
The data has been successfully extracted and chunked using **eparse**. Each data point maintains:
- **Row/Column Position**: Exact location in the original Excel file
- **Header Relationships**: Links to row and column headers
- **Data Type Information**: Python type inference
- **Sheet Context**: Source sheet information
- **Excel Reference**: Original cell reference (e.g., B10)

### 💡 Analysis Capabilities
- **Interactive Visualizations**: Explore data structure and distribution
- **Data Type Analysis**: Understand the variety of data types
- **Sheet-wise Analysis**: Compare data across different sheets
- **Column Analysis**: Examine data distribution by columns
- **Sample Data View**: Preview actual extracted data

### 🚀 Next Steps
- Use the visualizations to understand your data structure
- Export specific data subsets for further analysis
- Query the extracted data using the database interface
- Integrate with LLM tools for advanced data analysis
"""
    return summary

def download_enhanced_sample():
    """Create and return an enhanced sample Excel file for download."""
    filename = create_sample_excel()
    return filename

# Create Enhanced Gradio interface
def create_enhanced_interface():
    """Create the enhanced Gradio interface."""
    
    with gr.Blocks(title="Enhanced Excel Data Extraction & Visualization", theme=gr.themes.Soft()) as demo:
        gr.Markdown("# 🚀 Enhanced Excel Data Extraction & Visualization Tool")
        gr.Markdown("Upload an Excel file to extract, analyze, and visualize its data using advanced eparse integration")
        
        with gr.Row():
            with gr.Column(scale=1):
                file_input = gr.File(
                    label="Upload Excel File",
                    file_types=[".xlsx", ".xls"],
                    type="filepath"
                )
                
                with gr.Row():
                    upload_btn = gr.Button("📤 Process File", variant="primary", size="lg")
                    sample_btn = gr.Button("📥 Download Sample", variant="secondary")
                
                gr.Markdown("### 📋 Instructions")
                gr.Markdown("""
                1. **Upload** an Excel file (.xlsx or .xls)
                2. **Process** the file to extract data
                3. **Explore** the visualizations and summary
                4. **Download** a sample file to test the tool
                """)
                
                gr.Markdown("### 🔧 Features")
                gr.Markdown("""
                - **Advanced Data Extraction**: Full table detection
                - **Interactive Visualizations**: Multiple chart types
                - **Data Type Analysis**: Type inference and distribution
                - **Sheet-wise Analysis**: Per-sheet data breakdown
                - **Sample Data Preview**: View actual extracted data
                """)
            
            with gr.Column(scale=2):
                summary_output = gr.Markdown(label="Extraction Summary")
        
        with gr.Row():
            chart1 = gr.Plot(label="Data Overview", show_label=True)
            chart2 = gr.Plot(label="Sheets Distribution", show_label=True)
        
        with gr.Row():
            chart3 = gr.Plot(label="Columns Distribution", show_label=True)
            chart4 = gr.Plot(label="Data Types", show_label=True)
        
        with gr.Row():
            chart5 = gr.Plot(label="Sample Data", show_label=True)
            chart6 = gr.Plot(label="Data Summary", show_label=True)
        
        # Event handlers
        upload_btn.click(
            fn=process_excel_file_enhanced,
            inputs=[file_input],
            outputs=[summary_output, chart1, chart2, chart3, chart4, chart5, chart6]
        )
        
        sample_btn.click(
            fn=download_enhanced_sample,
            outputs=[gr.File(label="Enhanced Sample Excel File")]
        )
        
        gr.Markdown("---")
        gr.Markdown("### 🛠️ About This Enhanced Tool")
        gr.Markdown("""
        This enhanced tool provides comprehensive Excel data extraction and analysis:
        
        **🔍 Advanced Extraction:**
        - Full table detection and parsing
        - Header relationship preservation
        - Data type inference
        - Multi-sheet support
        
        **📊 Rich Visualizations:**
        - Interactive charts and graphs
        - Data distribution analysis
        - Type-based insights
        - Sample data preview
        
        **💾 Data Management:**
        - SQLite database storage
        - Structured data output
        - Export capabilities
        - Query interface
        
        Built with **Gradio**, **eparse**, and **Plotly** for powerful Excel data exploration.
        """)
    
    return demo

if __name__ == "__main__":
    # Create and launch the enhanced interface
    demo = create_enhanced_interface()
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=True,
        show_error=True
    ) 