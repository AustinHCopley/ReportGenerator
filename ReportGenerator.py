import openpyxl
from pandas import read_csv

# TODO: define class for generator, for now i will use a global array for id's
class Generator():
    """client sales report generator"""

    def __init__(self, csv='transactions.csv'):
        self.csv = csv
        self.xlsx = csv.split('.')[0] + '.xlsx'
        self.assets = []
        self.toXL()
        self.Total = self.wb.active
        
    def toXL(self):
        """convert csv to xlsx with pandas and open it as workbood with openpyxl"""
        try:
            self.transactions = read_csv(self.csv)
        except FileNotFoundError:
            print(f"File: {self.xlsx} not found")
            exit()
        self.transactions.to_excel(self.xlsx, sheet_name="Total", index=False)

        try:
            # create the instance's workbook
            self.wb = openpyxl.load_workbook(self.xlsx)
        except openpyxl.utils.exceptions.InvalidFileException:
            print(f"File: {self.xlsx} is not a valid excel file")
            exit()
    
    def setHeading(self, heading):
        self.heading = heading

    @staticmethod
    def insertRow(sheet, row, j = 1):
        for i, col in enumerate(row):
            sheet.cell(row=j, column=i+1).value = col.value



"""def insertRow(sheet, row, j = 1):
    #inserts row into sheet, j is the destination row
    for i, col in enumerate(row):
        sheet.cell(row=j, column=i+1).value = col.value"""


def main():

    # create generator instance
    gen = Generator("transactions.csv")
    
    # get the initial "Total" sheet from the original csv file
    

    # delete G-J, removing extraneous data that clients dont need to see
    gen.Total.delete_cols(7, 4)
    # save the header for copying into new sheets
    gen.setHeading(gen.Total[1])

    # Create separate sheets:
    # for each item in column B, if the assetID is new, create a new sheet
    for cell in gen.Total['B']:
        if cell.value != "Asset ID":
            if cell.value not in gen.assets:
                gen.assets.append(cell.value) # TODO: instead, append a tuple with (ID, name, price) to be sorted alphabetically and then based on price high to low
                gen.wb.create_sheet(title=f"p{gen.assets.index(cell.value)+1}")
                tempSheet = gen.wb[f"p{gen.assets.index(cell.value)+1}"]
                gen.insertRow(tempSheet, gen.heading)
                
    
    # add tab for vouchers and refunds
    vouchers = gen.wb.create_sheet(title="vouchers")
    gen.insertRow(vouchers, gen.heading)
    refunds = gen.wb.create_sheet(title="refunds")
    gen.insertRow(refunds, gen.heading)

    
    v = 2
    r = 2
    for row in gen.Total.iter_rows(min_row=2, max_row=gen.Total.max_row, min_col=1, max_col=19):
        # separate vouchers into separate sheet
        # all transactions with a value in "Note" column
        if row[6].value is not None:
            gen.insertRow(vouchers, row, v)
            v += 1

        # separate refunds into separate sheet
        # all remaining transactions without a value in "Referrer" column
        elif row[7].value is None:
            gen.insertRow(refunds, row, r)
            r += 1
        
        else:
            # TODO: seperate transactions according to assetID into proper sheet
            #       first need to get the proper order of assets alphabetically and then based on price high to low
            print(row[1].value) #show asset id
            print(gen.assets.index(row[1].value)) #show index of asset id

    gen.wb.save(gen.xlsx)

if __name__ == "__main__":
    main()
    print("Process complete")
    exit()