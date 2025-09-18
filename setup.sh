#!/bin/bash

echo "Excel GigaChat Analyzer - Setup Script"
echo "======================================"

# Check if uv is installed
if ! command -v uv &> /dev/null; then
    echo "Installing uv..."
    curl -LsSf https://astral.sh/uv/install.sh | sh
    echo "Please restart your shell or run 'source ~/.bashrc' to use uv"
    exit 1
fi

echo "✓ uv is installed"

# Install dependencies
echo "Installing dependencies..."
uv sync

if [ $? -eq 0 ]; then
    echo "✓ Dependencies installed successfully"
else
    echo "✗ Failed to install dependencies"
    exit 1
fi

# Check if .env exists
if [ ! -f ".env" ]; then
    echo "Creating .env file from template..."
    cp .env.example .env
    echo "✓ .env file created"
    echo "⚠ Please edit .env and add your GigaChat API key"
else
    echo "✓ .env file already exists"
fi

# Check if input.xlsx exists
if [ ! -f "input.xlsx" ]; then
    echo "⚠ input.xlsx not found - please add your Excel file to the project directory"
else
    echo "✓ input.xlsx found"
fi

echo ""
echo "Setup complete! Next steps:"
echo "1. Edit .env and add your GigaChat API key"
echo "2. Place your Excel file as input.xlsx (if not already done)"
echo "3. Run: uv run python main.py"
echo ""
echo "Or run the script directly: uv run excel-analyzer"
