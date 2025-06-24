#!/usr/bin/env python3
"""
Example script demonstrating Excel data extraction using eparse.
"""

import pandas as pd
import subprocess
import os
import tempfile
from pathlib import Path

def create_sample_excel():
    """Create a sample Excel file for testing."""
    data = {
        'Name': ['Alice Johnson', 'Bob Smith', 'Charlie Brown', 'Diana Prince'],
        'Age': [25, 30, 35, 28],
        'City': ['New York', 'Los Angeles', 'Chicago', 'Boston'],
        'Salary': [75000, 85000, 90000, 80000],
        'Department': ['Engineering', 'Marketing', 'Sales', 'HR']
    }
    
    df = pd.DataFrame(data)
    filename = 'sample_data.xlsx'
    df.to_excel(filename, index=False, sheet_name='Employees')
    
    # Create a second sheet with different data
    sales_data = {
        'Product': ['Laptop', 'Phone', 'Tablet', 'Monitor'],
        'Q1_Sales': [120, 200, 85, 45],
        'Q2_Sales': [135, 180, 95, 50],
        'Q3_Sales': [110, 220, 90, 40],
        'Q4_Sales': [125, 190, 100, 55]
    }
    
    with pd.ExcelWriter(filename, mode='a', if_sheet_exists='replace') as writer:
        pd.DataFrame(sales_data).to_excel(writer, sheet_name='Sales', index=False)
    
    print(f"Created sample Excel file: {filename}")
    return filename

def parse_excel_with_eparse(filename):
    """Parse Excel file using eparse command line tool."""
    try:
        # Parse with verbose output to see all tables
        result = subprocess.run([
            'eparse', '-f', filename, '-v', 'parse'
        ], capture_output=True, text=True, check=True)
        
        print("=== EPARSE VERBOSE OUTPUT ===")
        print(result.stdout)
        
        # Parse and output to console
        result = subprocess.run([
            'eparse', '-f', filename, '-o', 'stdout:///', 'parse'
        ], capture_output=True, text=True, check=True)
        
        print("=== EXTRACTED DATA ===")
        print(result.stdout)
        
    except subprocess.CalledProcessError as e:
        print(f"Error running eparse: {e}")
        print(f"Error output: {e.stderr}")

def parse_to_database(filename):
    """Parse Excel file and store in SQLite database."""
    try:
        # Create .files directory if it doesn't exist
        os.makedirs('.files', exist_ok=True)
        
        # Parse to SQLite database
        result = subprocess.run([
            'eparse', '-f', filename, '-o', 'sqlite3:///', 'parse', '-z'
        ], capture_output=True, text=True, check=True)
        
        print("=== DATABASE CREATION ===")
        print(result.stdout)
        
        # Find the created database file
        db_files = list(Path('.files').glob('*.db'))
        if db_files:
            db_file = db_files[0]
            print(f"Database created: {db_file}")
            
            # Query the database
            query_result = subprocess.run([
                'eparse', '-i', f'sqlite3:///{db_file}', '-o', 'stdout:///', 'query'
            ], capture_output=True, text=True, check=True)
            
            print("=== QUERY RESULTS ===")
            print(query_result.stdout)
            
    except subprocess.CalledProcessError as e:
        print(f"Error with database operations: {e}")
        print(f"Error output: {e.stderr}")

def main():
    """Main function to demonstrate eparse usage."""
    print("Excel Extraction Example using eparse")
    print("=" * 50)
    
    # Create sample Excel file
    filename = create_sample_excel()
    
    # Parse with eparse
    parse_excel_with_eparse(filename)
    
    # Parse to database
    parse_to_database(filename)
    
    # Clean up
    if os.path.exists(filename):
        os.remove(filename)
        print(f"\nCleaned up: {filename}")

if __name__ == "__main__":
    main() 