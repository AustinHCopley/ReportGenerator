import openpyxl
import pandas as pd


def main():
    csv = 'transactions.csv'
    xlsx = "test.xlsx"
    try:
        transactions = pd.read_csv(csv)
    except FileNotFoundError:
        print(f"File: {xlsx} not found")
        exit()
    transactions.to_excel(xlsx, sheet_name='Total', index=False)

    try:
        wb = openpyxl.load_workbook("./test.xlsx")
    except openpyxl.utils.exceptions.InvalidFileException:
        print(f"File: {xlsx} is not a valid excel file")
        exit()
    
    # get the initial sheet from the original csv file
    wsTotal = wb.active
    heading = wsTotal[1]

    # delete G-J, removing extraneous data that clients dont need to see
    wsTotal.delete_cols(7, 4)

    # create separate sheets:

    # for each item in column B, if the assetID is new, create a new sheet
    assets = []
    for cell in wsTotal['B']:
        if cell.value != "Asset ID":
            if cell.value not in assets:
                print(cell.value)
                assets.append(cell.value)
                wb.create_sheet(title=f"p{assets.index(cell.value)+1}")

    wb.save(xlsx)

    # add tab for vouchers and refunds
    wb.create_sheet(title="vouchers", index=-1)
    wb.create_sheet(title="refunds", index=-1)
    


# separate vouchers into separate sheet
# separate refunds into separate sheet


if __name__ == "__main__":
    main()
    print("Done")
    exit()