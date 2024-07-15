from datetime import date, timedelta
import win32com.client
import os
from os import listdir

### VARIABLES ###
### FUNCTIONS ###
def get_start_end_dates():
    day_of_week = date.today().weekday()  # Get today's day of the week as an index

    date_monday = date.today() - timedelta(
        days=day_of_week)

    str_date_Monday = date_monday.strftime(
        "%m-%d-%Y")

    #Calculate back to the previous Saturday
    date_Saturday = date.today() - timedelta(
        days=day_of_week + 2
    )  # Calculate the date of the Monday of this week
    str_date_Saturday = date_Saturday.strftime(
        "%m-%d-%Y"
    )  # Convert the complete date of Monday to a string
    return str_date_Saturday, str_date_Monday

def all_dir_paths():
    # Get the subdirectory name based on date
    day_of_week = date.today().weekday()  # Get today's day of the week as an index
    date_monday = date.today() - timedelta(
        days=day_of_week
    )  # Calculate the date of the Monday of this week
    str_date_monday = date_monday.strftime(
        "%Y-%m-%d"
    )  # Convert the complete date of Monday to a string

    dir_path = (
        # "C:\\Users\\darmstrong\\Desktop\\MSP-SecReview\\Weekly"
        "\\\\IMX-FS01\\MSP-SecReview\\weekly"  # Upper level directory for all client files
    )
    dir_list = listdir(dir_path)

    paths = []  # Create a list of all the current week directories
    for item in dir_list:
        if item[0].isdigit():
            paths.append(
                dir_path + "\\" + item + "\\" + str_date_monday
            )  # Add each file to the list as a full filepath
    return paths

def get_emails(dates, folders):
    saturday = dates[0]
    monday = dates[1]
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    ADReports = outlook.Folders['Security Reports'].Folders['Inbox'].Folders['Security Reports'].Folders['ADReports']

    dateFilter = f"[ReceivedTime] >= '{saturday}' AND [ReceivedTime] < '{monday}'"
    messages = ADReports.Items.Restrict(dateFilter)
    messages.Sort("[ReceivedTime]", True)

    # folders = all_dir_paths()

    for msg in messages:
        if str(msg.Subject).__contains__("0002"):
            directory = folders[0]
        elif str(msg.Subject).__contains__("0005"):
            directory = folders[1]
        elif str(msg.Subject).__contains__("0007"):
            directory = folders[2]
        elif str(msg.Subject).__contains__("0018"):
            directory = folders[3]
        elif str(msg.Subject).__contains__("0021"):
            directory = folders[4]
        elif str(msg.Subject).__contains__("0030"):
            directory = folders[5]
        elif str(msg.Subject).__contains__("0038"):
            directory = folders[6]
        elif str(msg.Subject).__contains__("0046"):
            directory = folders[7]
        elif str(msg.Subject).__contains__("0077"):
            directory = folders[8]
        elif str(msg.Subject).__contains__("0100"):
            directory = folders[9]
        elif str(msg.Subject).__contains__("0160"):
            directory = folders[10]
        elif str(msg.Subject).__contains__("0170"):
            directory = folders[11]
        elif str(msg.Subject).__contains__("171"):
            directory = folders[12]

        
            # print(msg.SenderName + " " + str(msg.ReceivedTime) + " " + str(msg.Subject))
            # Loop thru the attachments and save to file
        attachments = msg.Attachments
        i = 0
        for attachment in attachments:
            current_file = attachments[i]
            current_file.SaveAsFile(directory+'\\'+current_file.filename)
            i += 1

### Main ###
dates = get_start_end_dates()
folders = all_dir_paths()
get_emails(dates, folders)
### END ###