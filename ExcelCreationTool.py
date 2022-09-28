import random
from tkinter.font import BOLD
import xlsxwriter
import datetime
#cmd: pip3 install xlsxwriter

def main():

    # Creates formatted date variable.
    todaysDate = datetime.datetime.now()
    shortTodaysDate = todaysDate.strftime("%b%d")
    longTodaysDate = todaysDate.strftime("%x")

    # Variable for workbook name the user chooses.
    workbookName = input("Name of spreadsheet: ")
    workbookName = workbookName + "-" + shortTodaysDate

    # Data types to populate columns
    supplier = input("Enter Supplier: ").capitalize()
    product_model = input("Enter Product Model: ").capitalize()
    
    #TODO Add todays date for last inventoried on, formatted with datetime
    asset_status = ("Inventory")
    asset_location = ("PR-Bldg 70A - Information Technology (IT) (Room: 102)")
    asset_acquisition_date = ("1/2/2022")
    asset_expected_replacement_date = ("1/1/2026")
    asset_last_inventoried_on = (longTodaysDate)

     

    # Creates Excel workbook.
    workbook = xlsxwriter.Workbook(f"{workbookName}.xlsx")

    # Creates formatted date variable.
    todaysDate = datetime.datetime.now()
    todaysDate = todaysDate.strftime("%x")
    
    # Takes input from users, splits by whitespace, then creates list of split items.
    serials = list(map(str, input("Enter serial numbers with one space in between: ").split()))

    # Format for all cells below header.
    cellFormat = workbook.add_format()
    cellFormat.set_text_wrap()
    cellFormat.set_align("top")
    cellFormat.set_align("left")

    # Format for all cells in header.
    headerFormat = workbook.add_format()
    headerFormat.set_font(BOLD)
    headerFormat.set_text_wrap()
    headerFormat.set_align("left")

    # Creates a worksheet within the workbook.
    worksheet = workbook.add_worksheet("Assets")

    # Sets column width.
    worksheet.set_column("A:F", 20)
    worksheet.set_column("G:G", 50)
    worksheet.set_column("H:R", 20)
    

    # Writing to the header cells.
    worksheet.write("A1", "Serial Number", headerFormat) # Same as Service Tag and External ID
    worksheet.write("B1", "Supplier", headerFormat) # Will mostly be Apple, Dell, CDW, or Microsoft
    worksheet.write("C1", "Name", headerFormat) # Not Used
    worksheet.write("D1", "Product Model", headerFormat)
    worksheet.write("E1", "Status", headerFormat) # Defaults to "Inventory"
    worksheet.write("F1", "Service Tag", headerFormat) # Same as Serial Number and Service Tag
    worksheet.write("G1", "Location", headerFormat)
    worksheet.write("H1", "Owning Acct/Dept", headerFormat) # Not Used
    worksheet.write("I1", "Owner", headerFormat) # Not Used
    worksheet.write("J1", "Acquisition Date", headerFormat) # Default is "1/2/2022"
    worksheet.write("K1", "Maintenence Window", headerFormat)  # Not Used
    worksheet.write("L1", "Expected Replacement Date", headerFormat) # Default is "1/1/2026"
    worksheet.write("M1", "External ID", headerFormat) # Same as Serial Number and Service Tag
    worksheet.write("N1", "Notes", headerFormat) # Not Used
    worksheet.write("O1", "Last Inventoried On", headerFormat) # Use todays date in format 1/1/2022
    worksheet.write("P1", "Federal Capital #", headerFormat)  # Not Used
    worksheet.write("Q1", "Justification", headerFormat) # Not Used
    worksheet.write("R1", "PO Number", headerFormat) # Not Used
    

    

    # Starts the row at 2 to ignore the header row.
    rowIndex = 2

    # Defining a variable so the "serials" list iterator can be modified.
    cycleSerials = 0


    # Cycles through the length of the "serials" list to populate cells.
    for row in range(len(serials)):
        
        """ 
        Writes to the specified column, 
        the number of the row which gets inceremented by 1,
        the data type, then applies formatting
        """
        worksheet.write("A" + str(rowIndex), str(serials[cycleSerials]), cellFormat)
        worksheet.write("F" + str(rowIndex), str(serials[cycleSerials]), cellFormat)
        worksheet.write("M" + str(rowIndex), str(serials[cycleSerials]), cellFormat)
        worksheet.write("B" + str(rowIndex), supplier, cellFormat)
        worksheet.write("D" + str(rowIndex), product_model, cellFormat)
        worksheet.write("E" + str(rowIndex), asset_status, cellFormat)
        worksheet.write("G" + str(rowIndex), asset_location, cellFormat)
        worksheet.write("J" + str(rowIndex), asset_acquisition_date, cellFormat)
        worksheet.write("L" + str(rowIndex), asset_expected_replacement_date, cellFormat)
        worksheet.write("O" + str(rowIndex), asset_last_inventoried_on, cellFormat)

        # Outputs the serial numbers to the console, not integral, just visual.
        print(serials[cycleSerials])
        # Increments the "serial" list iterator by one to cycle through the list.
        cycleSerials = cycleSerials + 1
        # Increments "rowIndex" by one so the data can populate the cells downward
        rowIndex += 1

    workbook.close()
    

if __name__ == "__main__":
    main()
