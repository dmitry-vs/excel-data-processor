# Excel GigaChat Analyzer

A console application that reads an Excel file and uses the GigaChat API to calculate the average value in column C.

## Features

- Reads Excel files (.xlsx format)
- Extracts numeric data from column C
- Uses GigaChat API for calculation
- Provides manual verification of results
- Handles errors gracefully

## Prerequisites

- Python 3.8 or higher
- uv package manager
- GigaChat API key

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

## Configuration

1. Copy the environment template:
   ```bash
   cp .env.example .env
   ```

2. Edit `.env` and add your GigaChat API key:
   ```
   GIGA_KEY=your_actual_api_key_here
   ```

3. Place your Excel file as `input.xlsx` in the project directory

## Usage

Run the application:
```bash
uv run python main.py
```

Or using the script entry point:
```bash
uv run excel-analyzer
```

## Project Structure

```
.
├── main.py              # Main application
├── pyproject.toml       # Project configuration and dependencies
├── .env.example         # Environment variables template
├── .gitignore          # Git ignore rules
├── README.md           # This file
└── input.xlsx          # Your Excel file (not tracked by git)
```

## Dependencies

- `pandas`: Excel file reading and data manipulation
- `openpyxl`: Excel file format support
- `python-dotenv`: Environment variable management
- `gigachat`: GigaChat API client

## Error Handling

The application includes comprehensive error handling for:
- Missing environment variables
- File not found errors
- Invalid Excel file format
- Empty or non-numeric data in column C
- GigaChat API errors (with fallback to manual calculation)

## Notes

- The application automatically excludes `input.xlsx` and `.env` from git tracking
- If GigaChat API fails, the application falls back to manual calculation
- Only numeric values in column C are processed (non-numeric values are ignored)
