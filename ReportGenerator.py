import openpyxl
import pandas as pd



def main():
    csv = 'transactions.csv'
    xlsx = "test.xlsx"
    transactions = pd.read_csv(csv)
    transactions.to_excel(xlsx, sheet_name='Total', index=False)

    try:
        wb = openpyxl.load_workbook("./test.xlsx")
        sheet = wb.get_sheet_by_name('Total')
    except FileNotFoundError:
        print(f"File: {xlsx} not found")
        exit()
    except openpyxl.utils.exceptions.InvalidFileException:
        print(f"File: {xlsx} is not a valid excel file")
        exit()

# delete G-J
# separate vouchers into separate sheet
# separate refunds into separate sheet
#sheet.delete_cols(7, 4)

if __name__ == "__main__":
    main()
    print("Done")
    exit()