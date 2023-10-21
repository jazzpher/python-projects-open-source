import openpyxl
import sys

def main():
    print("This program writes data from a CSV file to an Excel file.")
    
    csv_name = input("Enter the name of the CSV file (with the extension): ")
    sep = input("Enter the separator used in the CSV file: ")
    excel_name = input("Enter the name of the Excel file (with the extension): ")
    sheet_name = input("Enter the name of the Excel sheet for output: ")

    try:
        wb = openpyxl.load_workbook(excel_name)
        sheet = wb[sheet_name]
        with open(csv_name, "r", encoding="utf-8") as file:
            write_data_to_excel(file, sheet, sep)
        wb.save(excel_name)
        print("Data has been successfully written to", excel_name)
    except Exception as e:
        print("Error:", str(e))
        sys.exit(1)

def write_data_to_excel(csv_file, excel_sheet, separator):
    row = 1

    for line in csv_file:
        line = line.rstrip('\n')
        data = line.split(separator)

        for column, value in enumerate(data, 1):
            excel_sheet.cell(row=row, column=column).value = value

        row += 1

if __name__ == "__main__":
    main()
