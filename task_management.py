import gspread
from constants import SCOPES,SHEET_NAME,SERVICE_ACCOUNT_FILE,COLORS
from datetime import datetime
from google.oauth2.service_account import Credentials
from gspread_formatting import *
from dateutil.parser import parse


class TaskManager:
    def __init__(client_setup):
        try: 
            credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        except Exception as e:
            print("Error setting JSON file:", e)
            exit()

        try:
            client_setup.client = gspread.authorize(credentials)
        except Exception as e:
            print("Authorization or sheet access problem:", e)
            exit()


        try:
                client_setup.spreadsheet = client_setup.client.open(SHEET_NAME)
                print("Sheet already exists. Using existing sheet.")
        except gspread.SpreadsheetNotFound:
             try:
                client_setup.spreadsheet = client_setup.client.create(SHEET_NAME)
                print("Sheet not found. Created new sheet:", client_setup.spreadsheet.url)
             except Exception as e:
                 print("Unable to create the Sheet Name",e)
                 exit()

             fmt = CellFormat(
                            horizontalAlignment='CENTER',
                            textFormat=TextFormat(
                            bold=True,
                            foregroundColor=COLORS["BLACK"]),
                            backgroundColor=COLORS["WHITE"],
                            borders=Borders(
                                           top=Border('SOLID', COLORS["BLACK"], 1),
                                            bottom=Border('SOLID', COLORS["BLACK"], 1),
                                            left=Border('SOLID', COLORS["BLACK"], 1),
                                            right=Border('SOLID', COLORS["BLACK"], 1)))
             sheetid = client_setup.spreadsheet.sheet1
             format_cell_range(sheetid, 'A1:C1', fmt)

             client_setup.spreadsheet.sheet1.insert_row(["Date", "Name", "Work"], index=1)

            
        client_setup.sheetid = client_setup.spreadsheet.sheet1

        print("Login Success")





    def get_sheet(sheet):
        return sheet.sheetid
    


    def get_information(task_input):
        try:
            name = input("Enter Name: ").strip()
            while not TaskManager.is_valid_name(name):
                name = input("Enter a valid name: ").strip()

            date = input("Enter Date: ").strip()
            str_date=TaskManager.parse_any_date(date)
            while not TaskManager.is_valid_date(str_date):
                date= input("Invalid date. Enter again : ").strip()
                str_date=TaskManager.parse_any_date(date)

            work = input("Work done: ").strip()
            while not TaskManager.is_valid_work(work):
                work = input("Work description too long. Enter again: ").strip()

            return str_date, name, work

        except KeyboardInterrupt:
            print("\nKeyboard interrupt detected. Exiting.")
            exit()

    
    def is_valid_name(name):
        return name.isalpha()

   
    def is_valid_work(work):
        return len(work) < 100
        #cell limit 50k

    def is_valid_date(date):
        try:
            datetime.strptime(date, "%Y-%m-%d")
            return True
        except:
            return False
        
    def has_today_entry(task_checker, name_to_check,date):
        try:
            records = task_checker.sheetid.get_all_records()
        except Exception as e:

            print("Unable to get the records of sheet")
            exit()

        for row in records:
            if row.get("Name") == name_to_check and row.get("Date") == date:
                return True
        return False
    
    def parse_any_date(date_str):
        try:
           dt = parse(date_str)
           entered_date = dt.date()
           current_date = datetime.today().date()
           if entered_date > current_date:
            print("Entry not possible with this date!!")
            exit()
           else:
           
            return entered_date.strftime("%Y-%m-%d")
        except ValueError:
              return None