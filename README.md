# Excel Data Processor

A console application that reads an Excel file and calculates statistics for column C data.

## Features

- Reads Excel files (.xlsx format)
- Extracts numeric data from column C
- Calculates average, sum, min, and max values
- Handles errors gracefully
- No external API dependencies

## Prerequisites

- Python 3.8 or higher
- uv package manager

## Installation

1. Install uv (if not already installed):
   ```bash
   curl -LsSf https://astral.sh/uv/install.sh | sh
   ```

2. Clone or download this project

3. Install dependencies:
   ```bash
   uv sync
   ```

## Usage

Place your Excel file as `input.xlsx` in the project directory, then run:

```bash
uv run python process_data.py
```

Or using the script entry point:
```bash
uv run data-processor
```

## Project Structure

```
.
├── process_data.py        # Main application
├── pyproject.toml         # Project configuration and dependencies
├── .gitignore            # Git ignore rules
├── README.md             # This file
└── input.xlsx            # Your Excel file (not tracked by git)
```

## Dependencies

- `pandas`: Excel file reading and data manipulation
- `openpyxl`: Excel file format support

## Error Handling

The application includes comprehensive error handling for:
- File not found errors
- Invalid Excel file format
- Empty or non-numeric data in column C

## Notes

- The application automatically excludes `input.xlsx` from git tracking
- Only numeric values in column C are processed (non-numeric values are ignored)
- No external API keys or configuration required