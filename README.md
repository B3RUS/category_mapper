# Sortownia Ofert v2.0 (Offer Sorter)

An automated product categorization tool that uses keyword-based rules to assign categories to offers/product listings.

## Features

- **Automated Categorization**: Assigns products to categories based on customizable keyword rules
- **GUI Application**: User-friendly Tkinter interface for easy operation
- **File Processing**: Supports batch processing of product listings from Excel/CSV files
- **Rule Management**: Define and manage keyword-to-category mappings
- **Category Configuration**: Centralized category definitions with IDs
- **Text Normalization**: Handles accents, special characters, and case-insensitive matching
- **Multi-language Support**: Can work with listings in multiple languages

## Project Structure

```
vibecode/
├── main.py                 # Main application with GUI
├── rules.json             # Keyword-to-category mapping rules
├── rules_example.json     # Example rules file
├── categories.json        # Category definitions and IDs
├── categories_example.json # Example categories file
└── README.md              # This file
```

## Files

### main.py
The main application file containing:
- GUI application class `AplikacjaKategorii` built with Tkinter
- Rule and category loading/saving functions
- Text normalization for reliable keyword matching
- File selection and processing interface

### rules.json
JSON array of `[keyword, category]` pairs. Example:
```json
[
  ["vintage gaming console", "Retro Electronics"],
  ["mechanical keyboard", "Computer Peripherals"],
  ["wireless headphones", "Audio Equipment"],
  ["mechanical mouse", "Computer Peripherals"],
  ["gaming monitor", "Display Equipment"],
  ["laptop stand", "Office Accessories"],
  ["desk lamp led", "Lighting"]
]
```

### categories.json
JSON object mapping category names to their IDs. Example:
```json
{
  "Electronics": 2001,
  "Retro Electronics": 2002,
  "Computer Peripherals": 2010,
  "Audio Equipment": 2020,
  "Display Equipment": 2030,
  "Office Accessories": 2040,
  "Lighting": 2050
}
```

## Requirements

- Python 3.10+
- pandas
- tkinter (usually included with Python)

## Installation

1. Clone or download the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python main.py
   ```

2. Configure your rules and categories:
   - Edit `rules.json` with keyword-to-category mappings
   - Edit `categories.json` with category definitions

3. In the GUI:
   - Select an input file (Excel/CSV with product listings)
   - Select an output file location
   - Click to process and categorize products

## How It Works

1. **Text Normalization**: Keywords are normalized to lowercase, accents removed, and special characters replaced with spaces for reliable matching
2. **Rule-Based Matching**: Products are categorized by finding matching keywords in the rules
3. **Batch Processing**: All products in the input file are processed and written to the output file with assigned categories

## Testing

Run the test suite with:
```bash
pytest
```

## Code Quality

The project includes automated linting with flake8:
```bash
flake8 . --count --exit-zero --max-complexity=10 --max-line-length=127
```

## License

[Add license information here]

## Contributing

[Add contribution guidelines here]
