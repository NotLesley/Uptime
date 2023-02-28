from download import Down
from upload import Up
import datetime
import raw_data_trans as raw
import time
import environ

# seconds
_PRO_DELAY = 30.0 
_DOWNLOAD_DELAY = 10.0 

env = environ.Env() 
environ.Env.read_env()

UPTIME_FOLDER = env('sharepoint_uptime_folder')
HEALTH_FOLDER = env('sharepoint_health_folder')
METRIC_FOLDER = env('sharepoint_metric_folder')
MONTHLY_FOLDER = env('sharepoint_monthly_folder')

root = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs"
monthly_root = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Monthly"

 # Get the current date and time
today = datetime.datetime.now().date()
_date = datetime.date.today()
date_string = today.strftime("%d %b %Y")
formatted_date = today.strftime("%Y_%m_%d")

# Split the formatted date into separate variables
day, month, year = date_string.split(" ")

# Calculate the date for the 3rd month from today
third_month_date = today - datetime.timedelta(days=90)

# Format the date as "MMMM" (e.g., "January", "February", "March", etc.)
third_month = third_month_date.strftime("%b")
third_month_year = third_month_date.strftime("%Y")

# Calculating yesterday's date
ytd = today - datetime.timedelta(days=1)
ytd_string = ytd.strftime("%d %b %Y")

# Source_file is the csv file from the database
source_file = rf"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\csv\uptime_{formatted_date}.csv"
daily_uptime = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Daily_Uptime_Report.xlsm"
weekly_uptime =  r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Weekly_Uptime_Report.xlsm"
health_report = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Health_Report.xlsx"
sishen_uptime = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Sishen_Uptime_Report.xlsm"
daily_metric = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\daily_matrix\Daily Metrics.xlsx"

def run():

    # Download workbook
    Down().download('None', UPTIME_FOLDER)
    Down().download("Daily Metrics.xlsx", METRIC_FOLDER)
    Down().download("Health_Report.xlsx", HEALTH_FOLDER)
    time.sleep(_DOWNLOAD_DELAY)
    print("Download complete")
    

    # Update Uptime
    raw.update(source_file, weekly_uptime, daily_uptime, health_report, sishen_uptime, daily_metric, date_string)
    time.sleep(_PRO_DELAY)

    # Creates monthly backups on the first weekday of a new month 
    if (0 <= today.weekday() <= 5 and today.day == 1) or (ytd.weekday() > 5 and ytd.day < 3):
        print("I am your Father")
        raw.saving(daily_uptime, third_month, third_month_year, 'Daily')
        raw.saving(weekly_uptime, third_month, third_month_year, 'Weekly')

        # Upload operating workbooks and monthly workbook
        Up().upload('Uptime', monthly_root, MONTHLY_FOLDER)
        Up().upload('Uptime', root, UPTIME_FOLDER)
    else:
        # Upload operating workbooks
        Up().upload('Uptime', root, UPTIME_FOLDER)

    Up().upload("Daily Metrics", root, METRIC_FOLDER)
    Up().upload("Health_Report", root, HEALTH_FOLDER)

if __name__ == '__main__':
    if _date.weekday() < 5:
        run()