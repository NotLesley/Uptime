from download import Down
from upload import Up
import datetime
import raw_data_trans as raw
import time

# seconds
_PRO_DELAY = 30.0 
_DOWNLOAD_DELAY = 10.0 

 # Get the current date and time
today = datetime.datetime.now().date()
now = datetime.datetime.now()
_date = datetime.date.today()
date_string = now.strftime("%d %b %Y")
formatted_date = today.strftime("%Y_%m_%d")

# Split the formatted date into separate variables
day, month, year = date_string.split(" ")

# Calculate the date for the 3rd month from today
third_month_date = now - datetime.timedelta(days=90)

# Format the date as "MMMM" (e.g., "January", "February", "March", etc.)
third_month = third_month_date.strftime("%b")

# Calculating yesterday's date
ytd = now - datetime.timedelta(days=4)
ytd_string = ytd.strftime("%d %b %Y")

# Source_file is the csv file from the database
source_file = rf"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\csv\uptime_{formatted_date}.csv"


def run():

    # Download workbook
    Down().download('None')
    print("Download complete")
    time.sleep(_DOWNLOAD_DELAY)

    daily_uptime = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Daily_Uptime_Report.xlsm"
    weekly_uptime =  r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Weekly_Uptime_Report.xlsm"

    # Update Uptime
    raw.update(source_file, weekly_uptime, daily_uptime, date_string, ytd_string)
    time.sleep(_PRO_DELAY)
   

    # creates monthly backups on the first weekday of a new month 
    if today.day == 1 and 0 < datetime.datetime.now().weekday() < 5:
        raw.saving(daily_uptime, third_month, month, year, 'daily')
        raw.saving(weekly_uptime, third_month, month, year, 'weekly')

        # Upload operating workbooks and monthly workbook
        Up().upload('uptime')
    else:
        # Upload operating workbooks
        Up().upload('Uptime')


if __name__ == '__main__':
    if _date.weekday() < 5:
        run()