from openpyxl import load_workbook, Workbook
from natsort import natsorted
import shutil
import glob
import re
import os
#asd
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


def ford():
    wb = load_workbook("1.FORD COURIER.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B2'] = cell.value
    ws1['C2'] = float(0.8333)

    wb1.save("calculation.xlsx")


def skoda():
    wb = load_workbook("2.SKODA FABIA.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B3'] = cell.value
    ws1['C3'] = float(0.7475)

    wb1.save("calculation.xlsx")


def skoda_2():
    wb = load_workbook("3.SKODA FABIA 2.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B4'] = cell.value
    ws1['C4'] = float(0.7475)

    wb1.save("calculation.xlsx")


def fiat():
    wb = load_workbook("4.FIAT PANDA.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B5'] = cell.value
    ws1['C5'] = float(0.7475)

    wb1.save("calculation.xlsx")


def citroen():
    wb = load_workbook("5.CITROEN JUMPER.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B6'] = cell.value
    ws1['C6'] = float(0.8333)

    wb1.save("calculation.xlsx")


def daewoo():
    wb = load_workbook("6.DAEWOO LUBLIN.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B7'] = cell.value
    ws1['C7'] = float(0.8333)

    wb1.save("calculation.xlsx")


def ford_2():
    wb = load_workbook("7.FORD TRANSIT.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B8'] = cell.value
    ws1['C8'] = float(0.8333)

    wb1.save("calculation.xlsx")


def ford_leasing():
    wb = load_workbook("8.FORD TRANSIT LEASING.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B9'] = cell.value
    ws1['C9'] = float(0.8333)

    wb1.save("calculation.xlsx")


def farmtrac():
    wb = load_workbook("9.FARMTRAC.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B10'] = cell.value
    ws1['C10'] = float(0.8333)

    wb1.save("calculation.xlsx")


def ursus():
    wb = load_workbook("10.URSUS.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B11'] = cell.value
    ws1['C11'] = float(0.8333)

    wb1.save("calculation.xlsx")


def tym():
    wb = load_workbook("11.TYM.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B12'] = cell.value
    ws1['C12'] = float(0.8333)

    wb1.save("calculation.xlsx")


def unimog():
    wb = load_workbook("12.UNIMOG.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B13'] = cell.value
    ws1['C13'] = float(0.8333)

    wb1.save("calculation.xlsx")


def unimog_2():
    wb = load_workbook("13.UNIMOG 2.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B14'] = cell.value
    ws1['C14'] = float(0.8333)

    wb1.save("calculation.xlsx")


def noremat():
    wb = load_workbook("14.NOREMAT.xlsx", data_only=True)
    ws = wb['ochrona_srodowiska']
    wb1 = load_workbook("calculation.xlsx")
    ws1 = wb1['Sheet']

    cell = ws['B13']
    ws1['B15'] = cell.value
    ws1['C15'] = float(0.8333)

    wb1.save("calculation.xlsx")


def cost():
    wb = load_workbook("calculation.xlsx")


# create_calculation()
# copy_files()
# create_file()
# sort_cars()
ford()
skoda()
skoda_2()
fiat()
citroen()
daewoo()
ford_2()
ford_leasing()
farmtrac()
ursus()
tym()
unimog()
unimog_2()
noremat()

# TODO: Dodaj if w celu wpisania aktualnej ceny
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
