from datetime import date, timedelta
import win32com.client
import os
from os import listdir

### VARIABLES ###
### FUNCTIONS ###
def get_start_end_dates():
    day_of_week = date.today().weekday()  # Get today's day of the week as an index
    date_monday = date.today() - timedelta(
        days=day_of_week -1)

    str_date_Monday = date_monday.strftime(
        "%m-%d-%Y")
    print(str_date_Monday)
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
    SentinelReports = outlook.Folders['Security Reports'].Folders['Inbox']

    dateFilter = f"[ReceivedTime] >= '{saturday}' AND [ReceivedTime] < '{monday}'"
    messages = SentinelReports.Items.Restrict(dateFilter)
    messages.Sort("[ReceivedTime]", True)

    # folders = all_dir_paths()

    for msg in messages:
        if str(msg.Subject).__contains__("0002") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[0]
        elif str(msg.Subject).__contains__("Knudtsen Chevrolet") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[1]
        elif str(msg.Subject).__contains__("bankcda") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[2]
        elif str(msg.Subject).__contains__("Coeur d'Alene Honda") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[3]
        elif str(msg.Subject).__contains__("Hospice of North Idaho") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[4]
        elif str(msg.Subject).__contains__("Magnuson McHugh Dougherty") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[5]
        elif str(msg.Subject).__contains__("Integrated Personnel") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[6]
        elif str(msg.Subject).__contains__("Northcon") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[7]
        elif str(msg.Subject).__contains__("Schweitzer Mountain Resort") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[8]
        elif str(msg.Subject).__contains__("Post Falls Family Medicine") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[9]
        elif str(msg.Subject).__contains__("ABRA Spokane") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[10]
        elif str(msg.Subject).__contains__("Knudtsen Foothills Mazda") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[11]
        elif str(msg.Subject).__contains__("Knudtsen Foothills ABRA") and str(msg.Subject).__contains__("Intermax Networks Reports"):
            directory = folders[12]
        else:
            continue
            # print(msg.SenderName + " " + str(msg.ReceivedTime) + " " + str(msg.Subject))
            # Loop thru the attachments and save to file
        try:
            attachments = msg.Attachments
            current_file = attachments[0]
            print(f"Moving {current_file.filename}")
            current_file.SaveAsFile(directory+'\\'+current_file.filename)
        except:
            pass

### Main ###
dates = get_start_end_dates()
folders = all_dir_paths()
get_emails(dates, folders)
### END ###