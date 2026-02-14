# Automated Budget Tools

Create professional Excel budget templates with Python. No manual formatting needed.

## What This Does

Automatically generates monthly budget spreadsheets with:
- ✅ Pre-built income and expense categories
- ✅ Automatic calculations (totals, balances, percentages)
- ✅ Professional formatting with color-coded inputs and formulas
- ✅ Currency formatting (zeros display as "-")
- ✅ Ready-to-use templates in seconds

## Quick Start

### Prerequisites

```bash
# Python 3.7+
python --version

# Setup Python Backend
cd backend-python
python -m venv venv
.\venv\Scripts\Activate.ps1  # Windows
pip install -r requirements.txt
```

### Generate Your Budget Template

```bash
# From the backend-python directory
python app.py
```

The Excel file will be created at `../outputs/budget-template/budget-template.xlsx`

## Project Structure

```
automated-budget-tools/
│
├── backend-python/
│   ├── app.py
│   ├── requirements.txt
│
├── backend-node/
│   ├── html2pptx-local.cjs
│   ├── package.json
│
├── public/
├── outputs/
├── README.md
```


## Dependencies

- **openpyxl** - Excel file creation and manipulation
- **Python 3.7+** - Core runtime

## Customization

Want to modify the template? Edit `outputs/budget-template/create_budget.py`:
- Add/remove budget categories
- Change color schemes
- Adjust formulas
- Modify number formatting

Then run the script again to generate your customized template.

## Technical Details

- **Zero formula errors** - All calculations use proper Excel formulas
- **Dynamic updates** - Change any input and totals recalculate automatically
- **No hardcoded values** - Everything uses cell references
- **Professional standards** - Follows financial modeling best practices

## License

This project is open source and available for personal and commercial use.

## Contributing

Contributions welcome! Feel free to:
- Add new budget categories
- Improve formatting
- Add new template types
- Enhance documentation

## Support

Found a bug or have a feature request? Open an issue on GitHub!
