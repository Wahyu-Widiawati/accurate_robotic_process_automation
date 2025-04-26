import fnmatch
import os
import pyautogui
import time
import logging
import subprocess
from datetime import datetime, timedelta

# Set the working directory
working_dir = r"C:\Users\USER.ADMIN01\Documents\Sales Invoice"
os.chdir(working_dir)

# Configure logging
logging.basicConfig(filename='rpa_log.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def log_in(password):
    try:
        logging.info('Starting the Accurate application.')
        subprocess.Popen(r"C:\Program Files (x86)\CPSSoft\ACCURATE5 Enterprise\Accurate.exe")
        pyautogui.press('enter')
        time.sleep(25)

        logging.info('Attempting to click on open_last_company.png.')
        pyautogui.click('open_last_company.png', r"C:\Users\USER.ADMIN01\Documents\Sales Invoice\reports.png")
        time.sleep(5)

        logging.info('Entering password.')
        pyautogui.write(password)


        pyautogui.press('enter')
        time.sleep(5)

        logging.info('Logged in successfully.')
    except Exception as e:
        logging.error(f'Error during login: {e}')

def find_button(path, confidence=0.8):
    button_location = pyautogui.locateCenterOnScreen(path, confidence=confidence)
    return button_location

def click_button(text, path, confidence=0.8):
    button_location = find_button(path, confidence)
    print(path)
    if button_location:
        pyautogui.click(button_location)
        logging.info(f"Clicked on button with text '{text}'")
    else:
        logging.error(f"Can't find button with text '{text}'")


def double_click_button(text, path, confidence=0.8):
    button_location = find_button(path, confidence)
    print(path)
    if button_location:
        pyautogui.doubleClick(button_location)
        logging.info(f"Double-clicked on button with text '{text}'")
    else:
        logging.error(f"Can't find button with text '{text}'")


def export_report(start_date, end_date):
    try:
        logging.info('Starting report export process.')

        # Open Reports
        logging.info('Clicking on reports.png.')
        click_button('Reports', 'reports.png')
        time.sleep(5)

        # Memorized Report
        logging.info('Clicking on memorized_report.png.')
        click_button('Memorized Report', r"C:\Users\USER.ADMIN01\Documents\Sales Invoice\memorized_report.png")
        time.sleep(5)

        # Sales Invoice_Tableau
        logging.info('Clicking on sales_invoice_tableau.png.')
        double_click_button('sales_invoice_tableau', 'sales_invoice_tableau.png')
        time.sleep(10)

        # Filter parameter date
        logging.info('Setting date filter.')
        double_click_button('From Date', 'from_date_button.png')
        pyautogui.press('delete')
        pyautogui.write(start_date)
        time.sleep(5)
        pyautogui.press('tab')
        pyautogui.write(end_date)
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(240) #600 for 5 days period

        # Export to Excel
        logging.info('Clicking on export.png.')
        click_button('Export', 'export.png')
        time.sleep(3)
        for i in range(4):
            pyautogui.press('tab')
        pyautogui.press('space')
        pyautogui.press('tab')
        pyautogui.press('space')
        for i in range(2):
            pyautogui.press('tab')
        pyautogui.press('space')
        for i in range(2):
            pyautogui.press('tab')
        pyautogui.press('space')
        pyautogui.press('enter')
        time.sleep(2)

        # Save the file
        folder_path = r"C:\Users\USER.ADMIN01\Documents\Sales Invoice"
        logging.info('Saving the file to %s', folder_path)
        pyautogui.press('delete')
        pyautogui.write(folder_path)
        (time.sleep(5))
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('delete')
        time.sleep(2)
        pyautogui.write(f'SI_{start_date}_{end_date}')
        time.sleep(5)
        pyautogui.press('enter')
        time.sleep(150)
        logging.info('Report exported successfully.')
    except Exception as e:
        logging.error(f'Error during report export: {e}')


def export_report_for_date_range(start_date, end_date):
    try:
        logging.info('Starting second report export process.')
        # Click modify
        logging.info(f'Try to click modify')
        click_button('modify', r"C:\Users\USER.ADMIN01\Documents\Sales Invoice\modify.png")
        time.sleep(7)
        logging.info('Setting date filter.')
        double_click_button('From Date', r"C:\Users\USER.ADMIN01\Documents\Sales Invoice\date_from_button.png")
        pyautogui.press('tab')
        pyautogui.press('delete')
        pyautogui.write(end_date)
        time.sleep(2)
        double_click_button('From Date', r"C:\Users\USER.ADMIN01\Documents\Sales Invoice\from_date_button.png")
        pyautogui.press('delete')
        pyautogui.write(start_date)
        time.sleep(2)
        pyautogui.press('enter')
        try:
            click_button('ok', r"C:\Users\USER.ADMIN01\Documents\Sales Invoicea\ok.png")
        except Exception as e:
            logging.error(f"Error occurred: {e}")
        time.sleep(240)  # 600 for 5 days period

        # Export to Excel
        logging.info('Clicking on export.png.')
        click_button('Export', r"C:\Users\USER.ADMIN01\Documents\Sales Invoice\export.png")
        time.sleep(3)
        pyautogui.press('enter')
        time.sleep(2)

        # Save the file
        logging.info('Saving the data')
        folder_path = r"C:\Users\USER.ADMIN01\Documents\Sales Invoice"
        logging.info('Saving the file to %s', folder_path)
        pyautogui.press('delete')
        pyautogui.write(folder_path)
        (time.sleep(5))
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('delete')
        time.sleep(2)
        pyautogui.write(f'SI_{start_date}_{end_date}')
        time.sleep(5)
        pyautogui.press('enter')
        try:
            click_button('Save', 'save.png')
        except Exception as e:
            logging.error(f"Error occurred: {e}")
        time.sleep(300)
        logging.info('Report exported successfully.')
    except Exception as e:
        logging.error(f'Error during report export: {e}')


# List files in folder
path = r"C:\Users\USER.ADMIN01\Documents\Sales Invoice"
list_all_files = os.listdir(path)
#print(list_all_files)

# Check the excel file only
list_excel_files = []
for file_name in list_all_files:
    if fnmatch.fnmatch(file_name, '*.xls'):
        list_excel_files.append(file_name)
        #print(file_name)

# List the available dates
## Split the dates based on the file names
list_available_dates = []
for file_name in list_excel_files:
    list_available_dates.append(file_name.split("_")[1]) if file_name else  None
    end_date_xls = file_name.split("_")[2]
    end_date = end_date_xls.split(".")[0]
    list_available_dates.append(end_date) if file_name else None
    #print(list_available_dates)
# Remove duplicates and sort the dates correctly
list_available_dates = sorted(set(list_available_dates), key=lambda date: datetime.strptime(date, "%d%m%Y"))

# Date ranges for available dates
available_date_ranges = []
for i in range(0, len(list_available_dates), 2):
    if i + 1 < len(list_available_dates):
        available_date_ranges.append((list_available_dates[i], list_available_dates[i + 1]))

#print(available_date_ranges)

## List dates in a year
def generate_date_ranges(start_date, end_date, delta_days=6):
    start = datetime.strptime(start_date, "%d%m%Y")
    end = datetime.strptime(end_date, "%d%m%Y")
    while start <= end:
        next_end = start + timedelta(days=delta_days)
        if next_end > end:
            next_end = end
        yield start.strftime("%d%m%Y"), next_end.strftime("%d%m%Y")
        start = next_end + timedelta(days=1)

# Main execution
password = '123'
log_in(password)
time.sleep(10)

start_date = "01012024"
end_date = "28022025"

date_ranges = list(generate_date_ranges(start_date, end_date))
# Find the missing date ranges
missing_date_ranges = [date_range for date_range in date_ranges if date_range not in available_date_ranges]
print(missing_date_ranges)

# Process the first range
export_report(missing_date_ranges[0][0], missing_date_ranges[0][1])
time.sleep(60)

# Process the subsequent ranges
for start, end in missing_date_ranges[1:]:
    export_report_for_date_range(start, end)
    time.sleep(10)
