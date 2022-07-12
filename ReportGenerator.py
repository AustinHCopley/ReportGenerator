import openpyxl
import pandas as pd

# TODO define class for generator, for now i will use a global array for id's
global assetOrder

def insertRow(sheet, row, j = 1):
    for i, col in enumerate(row):
        sheet.cell(row=j, column=i+1).value = col.value

def main():
    assetOrder = [3177496, 3127710, 3127729]
    csv = 'transactions.csv'
    xlsx = csv.split('.')[0] + '.xlsx'

    try:
        transactions = pd.read_csv(csv)
    except FileNotFoundError:
        print(f"File: {xlsx} not found")
        exit()
    transactions.to_excel(xlsx, sheet_name='Total', index=False)

    try:
        wb = openpyxl.load_workbook(xlsx)
    except openpyxl.utils.exceptions.InvalidFileException:
        print(f"File: {xlsx} is not a valid excel file")
        exit()
    
    # get the initial sheet from the original csv file
    wsTotal = wb.active

    # delete G-J, removing extraneous data that clients dont need to see
    wsTotal.delete_cols(7, 4)
    # save the header for copying into new sheets
    heading = wsTotal[1]

    # create separate sheets:

    # for each item in column B, if the assetID is new, create a new sheet
    assets = []
    for cell in wsTotal['B']:
        if cell.value != "Asset ID":
            if cell.value not in assets:
                assets.append(cell.value)
                wb.create_sheet(title=f"p{assets.index(cell.value)+1}")
                tempSheet = wb[f"p{assets.index(cell.value)+1}"]
                insertRow(tempSheet, heading)
                
    

    # add tab for vouchers and refunds
    vouchers = wb.create_sheet(title="vouchers")
    insertRow(vouchers, heading)
    refunds = wb.create_sheet(title="refunds")
    insertRow(refunds, heading)


# separate vouchers into separate sheet
    v = 2
    r = 2
    for row in wsTotal.iter_rows(min_row=2, max_row=wsTotal.max_row, min_col=1, max_col=19):
        if row[6].value is not None: #if row[0].value[:12] == "Used voucher":
            insertRow(vouchers, row, v)
            v += 1
        elif row[7].value is None:
            insertRow(refunds, row, r)
            r += 1
            


# separate refunds into separate sheet
    wb.save(xlsx)

if __name__ == "__main__":
    main()
    print("Done")
    exit()