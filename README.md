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

# Install dependencies
python -m venv venv
.\venv\Scripts\Activate.ps1  # Windows
pip install -r requirements.txt
```

### Generate Your Budget Template

```bash
# Activate virtual environment
.\venv\Scripts\Activate.ps1

# Run the generator
python outputs/budget-template/create_budget.py
```

The Excel file will be created at `outputs/budget-template/budget-template.xlsx`

## Features

### Budget Categories

**Income**
- Salary, Freelance/Side Income, Investment Income, Other Income

**Expenses**
- **Housing**: Rent/Mortgage, Utilities, Internet, Insurance, Maintenance
- **Transportation**: Car Payment, Gas, Insurance, Maintenance, Public Transit
- **Food**: Groceries, Dining Out, Coffee/Snacks
- **Personal**: Healthcare, Entertainment, Clothing, Personal Care, Subscriptions
- **Savings & Debt**: Emergency Fund, Retirement, Investments, Debt Payments
- **Other**: Gifts/Donations, Miscellaneous

### Automatic Calculations

- **Total Income** - Sums all income sources
- **Total Expenses** - Sums all expense categories
- **Net Balance** - Income minus Expenses
- **Savings Rate** - Percentage of income saved
- **Difference Column** - Budgeted vs Actual comparison

### Professional Formatting

Following industry-standard Excel conventions:
- 🔵 **Blue text** - User input cells (amounts you enter)
- ⚫ **Black text** - Formulas and calculations
- 🟡 **Yellow background** - Key assumptions (month/year)
- **Currency format**: `$#,##0;($#,##0);-` (zeros show as "-")
- **Percentages**: `0.0%` format

## How to Use the Template

1. **Open** the generated `budget-template.xlsx`
2. **Update Month/Year** in the yellow-highlighted cell
3. **Enter Budgeted Amounts** in column B (blue text)
4. **Track Actual Spending** in column C (blue text)
5. **Review** automatic calculations in totals and balance rows

All formulas update automatically as you enter data!

## Project Structure

```
automated-budget-tools/
├── outputs/
│   └── budget-template/
│       ├── budget-template.xlsx    # Generated Excel template
│       └── create_budget.py        # Generator script
├── requirements.txt                # Python dependencies
└── README.md                       # This file
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
