from task_management import TaskManager
from gspread_formatting import *
from constants import COLORS

fmt1 = CellFormat(
                    horizontalAlignment='LEFT',
                    textFormat=TextFormat(
                    bold=False,
                    foregroundColor=COLORS["BLACK"]),
                    backgroundColor=COLORS["WHITE"],
                    borders=Borders(top=Border('SOLID', COLORS["BLACK"], 1),
                                    bottom=Border('SOLID', COLORS["BLACK"], 1),
                                    left=Border('SOLID', COLORS["BLACK"], 1),
                                    right=Border('SOLID', COLORS["BLACK"], 1)))


tm_instance = TaskManager()
sheet = tm_instance.get_sheet()


choice = "Y"
while choice.upper() == "Y":

    date, name, work = tm_instance.get_information()
    if tm_instance.has_today_entry(name,date):
        print("User already entered todays task!!")
    else:
        sheet.append_row([date, name, work])
        last_row = len(sheet.get_all_values())
        format_cell_range(sheet, f'A{last_row}:C{last_row}', fmt1)
    

    try:
        choice = input("Do you want to continue adding? (Y/y): ")
    except KeyboardInterrupt:
        print("Keyboard Interrupt")
        exit()
