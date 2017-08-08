from openpyxl import Workbook

from writer import WorkSheetWriter


def write_to_excel(filename, table, position=(0, 0)):
    wb = Workbook()
    ws = wb.active
    WorkSheetWriter.write(ws, table, position)

    wb.save(filename)
