from openpyxl import *

# Check to see if the excel file exists (program was already run)
try:
    wb = load_workbook(filename= 'tool_checkout_system.xlsx')
except:
    # Create the workbook if it doesnt exist
    wb = Workbook()
    toolCheckoutLogSheet = wb.create_sheet("tool_checkout_log")
    employeesSheet = wb.create_sheet("employees")
    toolsSheet = wb.create_sheet("tools")
    toolCheckoutLogSheet['A1'] = "Employee"
    toolCheckoutLogSheet['B1'] = 'Tool'
    toolCheckoutLogSheet['C1'] = 'Sign Out Time'
    toolCheckoutLogSheet['D1'] = 'Sing In Time'

    employeesSheet['A1'] = 'Employee Number'
    employeesSheet['B1'] = 'Employee Name'

    toolsSheet['A1'] = 'Tool Number'
    toolsSheet['B1'] = 'Tool Name'
    wb.save('tool_checkout_system.xlsx')
    


