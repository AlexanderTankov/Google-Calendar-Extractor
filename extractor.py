import tkinter as tk

from datetime import datetime, timedelta
from tkcalendar import DateEntry

from excel_workbook import ExcelWorkbook
from calendar_extractor import CalendarExtractor
from Data.setup_credentials import setup_credentials


def extractor(start_date, end_date, workbook):
    '''TODO'''
    # Because you are getting only the date you have to convert it to datetime
    start_date = datetime.combine(start_date, datetime.min.time())
    end_date = datetime.combine(end_date, datetime.max.time())

    creds = setup_credentials()
    extractor = CalendarExtractor(start_date, end_date, creds)
    events = extractor.download_events()
    # TODO: Before store the events make sure that the events for the past 1 month are the same as in the excel
    num_stored_events = extractor.store_events_in_workbook(events, workbook)

    # Format the start and end date in order to looks better in the label msg
    start_date = start_date.strftime("%Y-%m-%d %H:%M")
    end_date = end_date.strftime("%Y-%m-%d %H:%M")

    result_lbl.config(text='You extracted {} events for the period of {} and {}'.format(num_stored_events,
                                                                                        start_date,
                                                                                        end_date))

def _start_date_entry_selected(event, end_date):
    '''Helper function for the start_date bind'''
    selected_date = event.widget.get_date()
    end_date.set_date(selected_date + timedelta(days=1))


if __name__ == '__main__':
    window = tk.Tk()
    window.title('Silvy Spa and Beauty')  
    window.geometry("600x400")

    # !IMPORTANT! If you change the name of the file do not forget to add it to .gitignore file
    workbook = ExcelWorkbook('calendar_events.xlsx')

    # If the file do not exist it will throw ValueError
    try:
        last_updated_date = datetime.strptime(workbook.get_last_updated_date(), '%Y-%m-%d %H:%M')
    except ValueError:
        last_updated_date = datetime.today()
    # As a start_date we will put the first day after the last update.
    first_not_updated_date = last_updated_date + timedelta(days=1)
    
    # If it is updated untill today we will not add this one day in adition to the starting date
    if last_updated_date.date() == datetime.today().date():
        first_not_updated_date = last_updated_date

    # Create date entrys for start and end date
    start_date = DateEntry(window, selectmode='day', year=first_not_updated_date.year,
                            month=first_not_updated_date.month, day=first_not_updated_date.day)
    start_date.grid(row =  2, column = 0, padx = 20, pady = 30)

    end_date = DateEntry(window, selectmode='day')
    end_date.grid(row = 2, column = 1, padx = 20, pady = 30)

    start_date.bind("<<DateEntrySelected>>", lambda event: _start_date_entry_selected(event, end_date))

    # Add button for execution
    exec_button = tk.Button(window, text = "Extract events",
             command = lambda: extractor(start_date.get_date(), 
                                         end_date.get_date(), 
                                         workbook))
    exec_button.grid(row = 2, column = 2, padx = 20, pady = 30)

    # Add label for the last updated date
    info_lbl = tk.Label(window, text = 'Last updated date is {}'.format(last_updated_date), wraplength=200)
    info_lbl.grid(row = 1, column = 0, padx = 20, pady = 30)
    
    # Add latel for the result of the extraction
    result_lbl = tk.Label(window, text = '', wraplength=200)
    result_lbl.grid(row = 1, column = 1, padx = 20, pady = 30)

    # Execute Tkinter
    window.mainloop()