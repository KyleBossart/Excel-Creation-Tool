
import xlsxwriter
import datetime
from my_lists import *
import tkinter
from tkinter import *
#cmd: pip3 install xlsxwriter tkinter


def main():
    master = tkinter.Tk()
    master.title("Asset Import")
    master.geometry("700x500")
    master.configure(bg="#2F567A")
    master.columnconfigure((0,1,2,3,4), weight=1)
    master.rowconfigure((0,1,2,3,4,5,6,7,8,9), weight=1)

    #Supplier List Dropdown Menu
    supplier_list_variable = StringVar(master)
    supplier_list_variable.set(supplier_list[0])
    supplier_list_dropdown = OptionMenu(master, supplier_list_variable, *supplier_list)
    supplier_list_dropdown.configure(bg="white", fg="black", width=40, font=12, borderwidth=0)

    # Product List Dropdown Menu
    product_list_variable = StringVar(master)
    product_list_variable.set(product_list[0])
    product_list_dropdown = OptionMenu(master, product_list_variable, *product_list)
    product_list_dropdown.configure(bg="white", fg="black", width=40, font=12, borderwidth=0)
    
    # Serial Description
    serials_description = tkinter.Label(master, text="Enter serial numbers below with a space between each one. \n Example: 10OC NM4D 0S51 31YK")
    serials_description.configure(bg="#2F567A", fg="white", font=12)

    # Serial Number Entry Box
    serials_entry = StringVar(master)
    serials_entry = tkinter.Entry(master, bg="white", fg="black", width=45, font=12, insertbackground="black")

    def submit():
        global supplier, product_model, serials
        supplier = (supplier_list_variable.get())
        product_model = (product_list_variable.get())
        serials = (serials_entry.get())
        master.destroy()
        

    submit_button = Button(master, text="Save and Submit", command=submit)
    submit_button.configure(bg="white", fg="black", borderwidth=0)

    footer = tkinter.Label(
        master, 
        text = "Kyle Bossart", 
        font=("Arial", 8), 
        bg="#2F567A", 
        fg="black"
        )
    
    
    supplier_list_dropdown.grid(row=1, column=2, pady=1)
    product_list_dropdown.grid(row=2, column=2, pady=1)
    serials_description.grid(row=4, column=2)
    serials_entry.grid(row=5, column=2)
    submit_button.grid(row=8, column=2, pady=5)
    footer.grid(row=9, column=0, sticky=SW, columnspan=4)

    master.mainloop()

    # Creates formatted date variable.
    todaysDate = datetime.datetime.now()
    full_date = todaysDate.strftime("%b-%d-%G-%f")
    last_inventory_date = todaysDate.strftime("%x")
    
    # Creates Excel file with name and date as the title.
    workbookName = ("Asset Import Spreadsheet - " + full_date)

    #TODO Add todays date for last inventoried on, formatted with datetime
    asset_status = ("Inventory")
    asset_location = ("PR-Bldg 70A - Information Technology (IT) (Room: 102)")
    asset_acquisition_date = ("1/2/2022")
    asset_expected_replacement_date = ("1/1/2026")
    asset_last_inventoried_on = (last_inventory_date)

     
    # Creates Excel workbook. In file path, use "\" for Windows, "/" for *nix.
    workbook = xlsxwriter.Workbook(f"Test Sheets\{workbookName}.xlsx")
    
    # Takes input from users, splits by whitespace, then creates list of split items.
    serials_list = list(map(str, (serials).split()))

    # Format for all cells below header.
    cellFormat = workbook.add_format()
    cellFormat.set_text_wrap()
    cellFormat.set_align("top")
    cellFormat.set_align("left")

    # Format for all cells in header.
    headerFormat = workbook.add_format()
    headerFormat.set_bold(True)
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
    for row in range(len(serials_list)):
        """ 
        Writes to the specified column, 
        the number of the row which gets inceremented by 1,
        the data type, then applies formatting
        """
        worksheet.write("A" + str(rowIndex), str(serials_list[cycleSerials]), cellFormat)
        worksheet.write("F" + str(rowIndex), str(serials_list[cycleSerials]), cellFormat)
        worksheet.write("M" + str(rowIndex), str(serials_list[cycleSerials]), cellFormat)
        worksheet.write("B" + str(rowIndex), supplier, cellFormat)
        worksheet.write("D" + str(rowIndex), product_model, cellFormat)
        worksheet.write("E" + str(rowIndex), asset_status, cellFormat)
        worksheet.write("G" + str(rowIndex), asset_location, cellFormat)
        worksheet.write("J" + str(rowIndex), asset_acquisition_date, cellFormat)
        worksheet.write("L" + str(rowIndex), asset_expected_replacement_date, cellFormat)
        worksheet.write("O" + str(rowIndex), asset_last_inventoried_on, cellFormat)

        # Outputs the serial numbers to the console, not integral, just visual.
        print(serials_list[cycleSerials])
        # Increments the "serial" list iterator by one to cycle through the list.
        cycleSerials = cycleSerials + 1
        # Increments "rowIndex" by one so the data can populate the cells downward
        rowIndex += 1
    
    workbook.close()
   

if __name__ == "__main__":
    main()
    
