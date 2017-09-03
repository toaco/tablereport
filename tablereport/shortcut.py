from openpyxl import Workbook

from .writer import WorkSheetWriter


def write_to_excel(filename, table, position=(0, 0)):
    """write table into excel. 
    
    If the file does not exist, a new file will be created. If the file has already
    existed, the file will be rewritten.
    
    By default, the table will be written in default worksheet at position (0,0).
    You can change the position by setting ``position`` argument.
    """
    wb = Workbook()
    ws = wb.active
    WorkSheetWriter.write(ws, table, position)

    wb.save(filename)
