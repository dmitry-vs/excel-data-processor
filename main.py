#!/usr/bin/env python3
"""
Excel GigaChat Analyzer

Console application that reads an Excel file and uses GigaChat API
to calculate the average value in column C.
"""

import os
import sys
import pandas as pd
from dotenv import load_dotenv
from gigachat import GigaChat
from gigachat.models import Chat, Messages, MessagesRole


def load_environment():
    """Load environment variables from .env file."""
    load_dotenv()
    api_key = os.getenv('GIGA_KEY')
    if not api_key:
        print("Error: GIGA_KEY not found in environment variables.")
        print("Please create a .env file with your GigaChat API key.")
        sys.exit(1)
    return api_key


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


def calculate_average_with_gigachat(data: list, api_key: str) -> float:
    """Use GigaChat to calculate the average of the given data."""
    try:
        # Initialize GigaChat client
        with GigaChat(credentials=api_key, verify_ssl_certs=False) as giga:
            # Prepare the data for the prompt
            data_str = ', '.join(map(str, data))
            
            prompt = f"""
            Please calculate the average (arithmetic mean) of the following numbers:
            {data_str}
            
            Please provide only the numerical result, rounded to 2 decimal places.
            """
            
            # Create chat request
            chat = Chat(
                messages=[
                    Messages(
                        role=MessagesRole.USER,
                        content=prompt
                    )
                ],
                temperature=0.1
            )
            
            # Get response
            response = giga.chat(chat)
            result_text = response.choices[0].message.content.strip()
            
            # Try to extract the numerical result
            try:
                # Look for a number in the response
                import re
                numbers = re.findall(r'-?\d+\.?\d*', result_text)
                if numbers:
                    return float(numbers[0])
                else:
                    # If no number found, try to parse the entire response
                    return float(result_text)
            except ValueError:
                print(f"Warning: Could not parse GigaChat response as number: {result_text}")
                # Fallback to manual calculation
                return sum(data) / len(data)
                
    except Exception as e:
        print(f"Error calling GigaChat API: {e}")
        print("Falling back to manual calculation...")
        return sum(data) / len(data)


def main():
    """Main application entry point."""
    print("Excel GigaChat Analyzer")
    print("=" * 30)
    
    # Load environment variables
    api_key = load_environment()
    
    # Read Excel file
    excel_file = "input.xlsx"
    df = read_excel_file(excel_file)
    
    # Get column C data
    column_c_data = get_column_c_data(df)
    
    print(f"\nData from column C: {column_c_data[:10]}{'...' if len(column_c_data) > 10 else ''}")
    
    # Calculate average using GigaChat
    print("\nCalculating average using GigaChat...")
    average = calculate_average_with_gigachat(column_c_data, api_key)
    
    # Display results
    print(f"\nResults:")
    print(f"Number of values: {len(column_c_data)}")
    print(f"Average value: {average:.2f}")
    
    # Verify with manual calculation
    manual_average = sum(column_c_data) / len(column_c_data)
    print(f"Manual verification: {manual_average:.2f}")
    
    if abs(average - manual_average) < 0.01:
        print("✓ GigaChat calculation matches manual calculation")
    else:
        print("⚠ GigaChat calculation differs from manual calculation")


if __name__ == "__main__":
    main()
