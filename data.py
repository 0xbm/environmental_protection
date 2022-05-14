from openpyxl import load_workbook, Workbook
from natsort import natsorted
import shutil
import glob
import re
import os
import cars_list as c

# TODO:ZROBIC KLASE NA PODSTWIE DOSTEPNYCH FUNKCJI
# class create_file():
def create_file():
    wb = Workbook()
    ws = wb['Sheet']
    ws['A1'] = 'POJAZDY'
    ws['B1'] = 'ILOSC PALIWA'
    ws['C1'] = 'WAGA PALIWA'
    ws['D1'] = 'CENA'
    # for row in range(1):
    #    ws1.append(range(0, 13))

    wb.save("calculation.xlsx")


def copy_files():
    source_dir = "/Users/btn/PycharmProjects/cars/2022/"
    dest_dir = "/Users/btn/PycharmProjects/environmental_protection/"
    files = glob.iglob(os.path.join(source_dir, "*.xlsx"))
    for file in files:
        if os.path.isfile(file):
            shutil.copy2(file, dest_dir)


def sort_cars():
    path = "/Users/btn/PycharmProjects/environmental_protection/"
    files = [f for f in glob.glob(path + "**/*.xlsx", recursive=True)]
    folders_show = re.compile(r".*(/environmental_protection/)")
    files_sort = natsorted([folders_show.sub('', x).strip() for x in files])
    print(files_sort)

    wb = load_workbook("calculation.xlsx")
    ws = wb['Sheet']
    ws.column_dimensions["A"].width = 25
    ws['A2'] = files_sort[0]
    ws['A3'] = files_sort[1]
    ws['A4'] = files_sort[2]
    ws['A5'] = files_sort[3]
    ws['A6'] = files_sort[4]
    ws['A7'] = files_sort[5]
    ws['A8'] = files_sort[6]
    ws['A9'] = files_sort[7]
    ws['A10'] = files_sort[8]
    ws['A11'] = files_sort[9]
    ws['A12'] = files_sort[10]
    ws['A13'] = files_sort[11]
    ws['A14'] = files_sort[12]
    ws['A15'] = files_sort[13]

    ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 8
    wb.save("calculation.xlsx")

def cost():
    wb = load_workbook("calculation.xlsx")
    ws = wb['Sheet']

    ford = ws['A2'].value
    fc = int(input("Cost for " + ford + ": "))
    ws['D2'] = fc

    skoda = ws['A3'].value
    sf = int(input("Cost for " + skoda + ": "))
    ws['D3'] = sf





    wb.save("calculation.xlsx")


# create_calculation()  #working
# copy_files()          #working
# create_file()         #working
# sort_cars()           #working
cost()                # TODO: Sprzawdz w ktorych latach sa dane samochody i dokoncz progsa


'''
c.ford()
c.skoda()
c.skoda_2()
c.fiat()
c.citroen()
c.daewoo()
c.ford_2()
c.ford_leasing()
c.farmtrac()
c.ursus()
c.tym()
c.unimog()
c.unimog_2()
c.noremat()
'''


'''
def copy_datas():
    # z listy skopiowac nazwy pojazdow do komorek
    wb = load_workbook("1.FORD COURIER.xlsx")
    sheet = wb["ochrona_srodowiska"]

    wb1 = load_workbook("calculation.xlsx")
    sheet1 = wb1["Sheet1"]

    # Copy range of cells as a nested list
    # Takes: start cell, end cell, and sheet you want to copy from.
    def copyRange(startCol, startRow, endCol, endRow, sheet):
        rangeSelected = []
        # Loops through selected Rows
        for i in range(startRow, endRow + 1, 1):
            # Appends the row to a RowSelected list
            rowSelected = []
            for j in range(startCol, endCol + 1, 1):
                rowSelected.append(sheet.cell(row=i, column=j).value)
            # Adds the RowSelected List and nests inside the rangeSelected
            rangeSelected.append(rowSelected)

        return rangeSelected

    # Paste range
    # Paste data from copyRange into template sheet
    def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving, copiedData):
        countRow = 0
        for i in range(startRow, endRow + 1, 1):
            countCol = 0
            for j in range(startCol, endCol + 1, 1):
                sheetReceiving.cell(row=i, column=j).value = copiedData[countRow][countCol]
                countCol += 1
            countRow += 1

    def createData():
        print("Processing...")
        selectedRange = copyRange(2, 13, 2, 13, sheet)  # Change the 4 number values
        pastingRange = pasteRange(1, 3, 4, 15, sheet1, selectedRange)  # Change the 4 number values
        # You can save the template as another file to create a new file here too.s
        wb1.save("calculation.xlsx")
'''
