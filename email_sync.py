import openpyxl
import shutil
import os




def build_queries():
        
    wb = openpyxl.load_workbook(filename='Email_updates.xlsx', read_only=False)
    sheet = wb.get_sheet_by_name('EMAILS')
    writeQueries = {}
    for row in range(2, sheet.max_row + 1):
        EMPL_ID = sheet['A' + str(row)].value
        BUSN_EMAIL = sheet['C' + str(row)].value
        PERS_EMAIL = sheet['D' + str(row)].value
        print(EMPL_ID, BUSN_EMAIL, PERS_EMAIL)

       



def main():
    os.chdir('H:\SavedQueries')
    build_queries()
    


if __name__ == "__main__":
    main()
