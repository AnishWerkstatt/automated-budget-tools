from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()
sheet = wb.active
sheet.title = "Monthly Budget"

# Color definitions (per XLSX skill standards)
BLUE = "0000FF"  # Inputs
BLACK = "000000"  # Formulas
YELLOW = "FFFF00"  # Key assumptions
HEADER_FILL = "D3D3D3"  # Light gray for headers

# Fonts
input_font = Font(color=BLUE, bold=False)
formula_font = Font(color=BLACK, bold=False)
header_font = Font(bold=True, size=12)
section_font = Font(bold=True, size=11)

# Fills
yellow_fill = PatternFill(start_color=YELLOW, end_color=YELLOW, fill_type="solid")
header_fill = PatternFill(start_color=HEADER_FILL, end_color=HEADER_FILL, fill_type="solid")

# Border
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

# Title
sheet['A1'] = 'Monthly Budget Template'
sheet['A1'].font = Font(bold=True, size=14)
sheet.merge_cells('A1:D1')

# Month/Year
sheet['A2'] = 'Month:'
sheet['B2'] = 'January 2024'
sheet['B2'].font = input_font
sheet['B2'].fill = yellow_fill

# Column headers
sheet['A4'] = 'Category'
sheet['B4'] = 'Budgeted'
sheet['C4'] = 'Actual'
sheet['D4'] = 'Difference'
for col in ['A', 'B', 'C', 'D']:
    sheet[f'{col}4'].font = header_font
    sheet[f'{col}4'].fill = header_fill
    sheet[f'{col}4'].alignment = Alignment(horizontal='center')

# INCOME SECTION
row = 5
sheet[f'A{row}'] = 'INCOME'
sheet[f'A{row}'].font = section_font

income_categories = ['Salary', 'Freelance/Side Income', 'Investment Income', 'Other Income']
income_start = row + 1
for cat in income_categories:
    row += 1
    sheet[f'A{row}'] = cat
    sheet[f'B{row}'].font = input_font
    sheet[f'C{row}'].font = input_font

income_end = row

# Total Income
row += 1
sheet[f'A{row}'] = 'Total Income'
sheet[f'A{row}'].font = Font(bold=True)
sheet[f'B{row}'] = f'=SUM(B{income_start}:B{income_end})'
sheet[f'B{row}'].font = formula_font
sheet[f'C{row}'] = f'=SUM(C{income_start}:C{income_end})'
sheet[f'C{row}'].font = formula_font
sheet[f'D{row}'] = f'=C{row}-B{row}'
sheet[f'D{row}'].font = formula_font
total_income_row = row

# EXPENSES SECTION
row += 2
sheet[f'A{row}'] = 'EXPENSES'
sheet[f'A{row}'].font = section_font

# Housing
row += 1
sheet[f'A{row}'] = 'Housing'
sheet[f'A{row}'].font = Font(bold=True, italic=True)
housing_categories = ['Rent/Mortgage', 'Utilities (Electric, Gas, Water)', 'Internet/Cable', 'Home Insurance', 'Maintenance/Repairs']
housing_start = row + 1
for cat in housing_categories:
    row += 1
    sheet[f'A{row}'] = f'  {cat}'
    sheet[f'B{row}'].font = input_font
    sheet[f'C{row}'].font = input_font
housing_end = row

# Transportation
row += 1
sheet[f'A{row}'] = 'Transportation'
sheet[f'A{row}'].font = Font(bold=True, italic=True)
transport_categories = ['Car Payment', 'Gas/Fuel', 'Car Insurance', 'Maintenance/Repairs', 'Public Transit']
transport_start = row + 1
for cat in transport_categories:
    row += 1
    sheet[f'A{row}'] = f'  {cat}'
    sheet[f'B{row}'].font = input_font
    sheet[f'C{row}'].font = input_font
transport_end = row

# Food
row += 1
sheet[f'A{row}'] = 'Food'
sheet[f'A{row}'].font = Font(bold=True, italic=True)
food_categories = ['Groceries', 'Dining Out', 'Coffee/Snacks']
food_start = row + 1
for cat in food_categories:
    row += 1
    sheet[f'A{row}'] = f'  {cat}'
    sheet[f'B{row}'].font = input_font
    sheet[f'C{row}'].font = input_font
food_end = row

# Personal
row += 1
sheet[f'A{row}'] = 'Personal'
sheet[f'A{row}'].font = Font(bold=True, italic=True)
personal_categories = ['Healthcare/Medical', 'Entertainment', 'Clothing', 'Personal Care', 'Subscriptions']
personal_start = row + 1
for cat in personal_categories:
    row += 1
    sheet[f'A{row}'] = f'  {cat}'
    sheet[f'B{row}'].font = input_font
    sheet[f'C{row}'].font = input_font
personal_end = row

# Savings & Debt
row += 1
sheet[f'A{row}'] = 'Savings & Debt'
sheet[f'A{row}'].font = Font(bold=True, italic=True)
savings_categories = ['Emergency Fund', 'Retirement Savings', 'Investment Contributions', 'Debt Payments']
savings_start = row + 1
for cat in savings_categories:
    row += 1
    sheet[f'A{row}'] = f'  {cat}'
    sheet[f'B{row}'].font = input_font
    sheet[f'C{row}'].font = input_font
savings_end = row

# Other
row += 1
sheet[f'A{row}'] = 'Other'
sheet[f'A{row}'].font = Font(bold=True, italic=True)
other_categories = ['Gifts/Donations', 'Miscellaneous']
other_start = row + 1
for cat in other_categories:
    row += 1
    sheet[f'A{row}'] = f'  {cat}'
    sheet[f'B{row}'].font = input_font
    sheet[f'C{row}'].font = input_font
other_end = row

# Total Expenses
row += 1
expenses_end = row - 1
sheet[f'A{row}'] = 'Total Expenses'
sheet[f'A{row}'].font = Font(bold=True)
sheet[f'B{row}'] = f'=SUM(B{housing_start}:B{other_end})'
sheet[f'B{row}'].font = formula_font
sheet[f'C{row}'] = f'=SUM(C{housing_start}:C{other_end})'
sheet[f'C{row}'].font = formula_font
sheet[f'D{row}'] = f'=C{row}-B{row}'
sheet[f'D{row}'].font = formula_font
total_expenses_row = row

# NET BALANCE
row += 2
sheet[f'A{row}'] = 'NET BALANCE (Income - Expenses)'
sheet[f'A{row}'].font = Font(bold=True, size=12)
sheet[f'B{row}'] = f'=B{total_income_row}-B{total_expenses_row}'
sheet[f'B{row}'].font = Font(color=BLACK, bold=True)
sheet[f'C{row}'] = f'=C{total_income_row}-C{total_expenses_row}'
sheet[f'C{row}'].font = Font(color=BLACK, bold=True)
sheet[f'D{row}'] = f'=C{row}-B{row}'
sheet[f'D{row}'].font = Font(color=BLACK, bold=True)
net_balance_row = row

# Percentage Analysis
row += 2
sheet[f'A{row}'] = 'BUDGET ANALYSIS'
sheet[f'A{row}'].font = section_font

row += 1
sheet[f'A{row}'] = 'Savings Rate (% of Income)'
sheet[f'B{row}'] = f'=IF(B{total_income_row}=0,0,B{net_balance_row}/B{total_income_row})'
sheet[f'B{row}'].font = formula_font
sheet[f'C{row}'] = f'=IF(C{total_income_row}=0,0,C{net_balance_row}/C{total_income_row})'
sheet[f'C{row}'].font = formula_font

# Add difference formulas for all expense rows
for r in range(income_start, other_end + 1):
    if sheet[f'B{r}'].value is None:
        sheet[f'D{r}'] = f'=C{r}-B{r}'
        sheet[f'D{r}'].font = formula_font

# Number formatting
currency_format = '$#,##0;($#,##0);-'
percentage_format = '0.0%'

for r in range(income_start, net_balance_row + 1):
    for col in ['B', 'C', 'D']:
        if sheet[f'{col}{r}'].value:
            sheet[f'{col}{r}'].number_format = currency_format

# Format percentage cells
sheet[f'B{row}'].number_format = percentage_format
sheet[f'C{row}'].number_format = percentage_format

# Column widths
sheet.column_dimensions['A'].width = 35
sheet.column_dimensions['B'].width = 15
sheet.column_dimensions['C'].width = 15
sheet.column_dimensions['D'].width = 15

# Save
wb.save('../outputs/budget-template/budget-template.xlsx')
print("Budget template created successfully!")
