#!/usr/bin/env python3
"""
Excel Data Processor

Console application that reads an Excel file and calculates
the average value in column C.
"""

import sys
import pandas as pd


def read_excel_file(file_path: str) -> pd.DataFrame:
    """Read Excel file and return DataFrame."""
    try:
        df = pd.read_excel(file_path)
        print(f"Successfully loaded Excel file: {file_path}")
        print(f"Shape: {df.shape[0]} rows, {df.shape[1]} columns")
        return df
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)


def get_column_c_data(df: pd.DataFrame) -> list:
    """Extract data from column C (index 2)."""
    if df.shape[1] < 3:
        print("Error: Excel file doesn't have enough columns (need at least 3 for column C).")
        sys.exit(1)
    
    column_c = df.iloc[:, 2]  # Column C is index 2
    # Convert to numeric, ignoring non-numeric values
    numeric_data = pd.to_numeric(column_c, errors='coerce')
    numeric_data = numeric_data.dropna()  # Remove NaN values
    
    if len(numeric_data) == 0:
        print("Error: No numeric data found in column C.")
        sys.exit(1)
    
    print(f"Found {len(numeric_data)} numeric values in column C")
    return numeric_data.tolist()


def calculate_average(data: list) -> float:
    """Calculate the average of the given data."""
    if not data:
        return 0.0
    return sum(data) / len(data)


def main():
    """Main application entry point."""
    print("Excel Data Processor")
    print("=" * 30)
    
    # Read Excel file
    excel_file = "input.xlsx"
    df = read_excel_file(excel_file)
    
    # Get column C data
    column_c_data = get_column_c_data(df)
    
    print(f"\nData from column C: {column_c_data[:10]}{'...' if len(column_c_data) > 10 else ''}")
    
    # Calculate average
    print("\nCalculating average...")
    average = calculate_average(column_c_data)
    
    # Display results
    print(f"\nResults:")
    print(f"Number of values: {len(column_c_data)}")
    print(f"Average value: {average:.2f}")
    print(f"Sum: {sum(column_c_data):.2f}")
    print(f"Min: {min(column_c_data):.2f}")
    print(f"Max: {max(column_c_data):.2f}")


if __name__ == "__main__":
    main()
