from openpyxl import *

try:
    wb = load_workbook(filename= 'tool_checkout_system.xlsx')
    tool_sheet = wb['tools']
    employee_sheet = wb['employees']
except:
    print("Error loading workbook.")
    print("Make sure to run checkout_system before running this file")


def enterData(number, name, sheet):
        dataToAppend = [number,name]
        sheet.append(dataToAppend)
        wb.save('tool_checkout_system.xlsx')



while True:
    print("\nWelcome to the tool checkout system")
    print("[DATA ENTRY MODE]")
    print("------------------------------------")
    # Collect and Validate input
    isValidInput = False
    while not isValidInput:
        action = input("Enter 1 to ENTER A TOOL, 2 to ENTER AN EMPLOYEE, or 'q' to QUIT: ")
        if action not in ['1','2','q']:
            print("\n*** ERROR: Invalid Input. ***\n")
        else:
            isValidInput = True
            if action == 'q':
                wb.save('tool_checkout_system.xlsx')
                wb.close()
                quit()            

    match action:
        case "1":
            toolNumber = int(input("Scan the tool barcode: "))
            toolName = input("Enter the tool name: ")
            enterData(toolNumber,toolName,tool_sheet)
            print(f"\nTOOL #{toolNumber} ENTERED AS {toolName}\n")

        case "2": 
            employeeNumber = int(input("Scan your employee badge: "))
            employeeName = input("Enter the employee name: ")
            enterData(employeeNumber,employeeName,employee_sheet)
            print(f"\nEMPLOYEE #{employeeNumber} ENTERED AS {employeeName}\n")




    