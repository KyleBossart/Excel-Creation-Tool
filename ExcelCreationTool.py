import random
import xlsxwriter
import datetime


def main():

    # Creates formatted date variable.
    todaysDate = datetime.datetime.now()
    shortTodaysDate = todaysDate.strftime("%b%d%m")
    longTodaysDate = todaysDate.strftime("%x")

    # Variable for workbook name the user chooses.
    workbookName = input("Name of spreadsheet: ")
    workbookName = workbookName + "-" + shortTodaysDate

    # Data types to populate columns
    make = input("Enter Make: ").capitalize()
    model = input("Enter Model: ").capitalize()

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
    headerFormat.set_bg_color("#0AB8F4")
    headerFormat.set_text_wrap()
    headerFormat.set_align("left")

    # Creates a worksheet within the workbook.
    worksheet = workbook.add_worksheet("Assets")

    # Sets column width.
    worksheet.set_column("A:D", 20)

    # Writing to the header cells.
    worksheet.write("A1", "Make", headerFormat)
    worksheet.write("B1", "Model", headerFormat)
    worksheet.write("C1", "Serial Number", headerFormat)
    worksheet.write("D1", "Date Imported", headerFormat)

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
        worksheet.write("A" + str(rowIndex), make, cellFormat)
        worksheet.write("B" + str(rowIndex), model, cellFormat)
        worksheet.write("C" + str(rowIndex), str(serials[cycleSerials]), cellFormat)
        worksheet.write("D" + str(rowIndex), longTodaysDate, cellFormat)

        # Outputs the data to the console, not integral, just visual.
        print(make, model, serials[cycleSerials])
        # Increments the "serial" list iterator by one to cycle through the list.
        cycleSerials = cycleSerials + 1
        # Increments "rowIndex" by one so the data can populate the cells downward
        rowIndex += 1

    workbook.close()
    

if __name__ == "__main__":
    main()
