import PyPDF2
from PyPDF2 import PdfFileMerger
import os
from os import listdir
from win32com import client
from datetime import date, timedelta
import shutil

# import time

##########################
# DICTIONARIES AND LISTS #
##########################

# Dictionary of keywords for the pdf files
f_key = {
    "app_use_sum": "Application_Control_Application_Usage_Summary",
    "app_cont": "Application_Control_Blocked_Application_Summary",
    "app_use_bw": [
        "Application_Usage_Applications_(Host)_by_Bandwidth",
        "Application_Usage_Applications__Host__by_Bandwidth",
        "ALL_Application_Usage_Applications__User__by_Bandwidth",
    ],
    "app_use_hits": [
        "Application_Usage_Applications_(Host)_by_Hits",
        "Application_Usage_Applications__Host__by_Hits",
        "Application_Usage_Top_Hosts_by_Hits",
    ],
    "block_sites": [
        "Blocked_Websites_Category",
        "Blocked_Websites_Client",
        "ALL_Blocked_Websites_Category",
    ],
    "botnets": [
        "Botnet_Detection_Activity_Trend",
        "BotNet Detection Details",
        "Blocked_Botnet_Sites",
    ],
    "botnet_detect": [
        "Botnet_Detection_Botnet_Detection_by_Client",
        "Botnet_Detection_Botnet_Detection_by_Destination",
    ],
    "IPS": [
        "Intrusions_(IPS)_Activity_Trend",
        "Intrusions__IPS__Activity_Trend",
        "Intrusions__IPS__Source",
        "Intrusions__IPS__Signatures",
        "Intrusions__IPS__Protocol",
        "Intrusions__IPS__Threat_Level",
        "Intrusions Prevention Details",
        "IPS Details.pdf",
        "ALL_Intrusions__IPS__Activity_Trend",
    ],
    "active_client": [
        "Most_Active_Clients_Clients_by_Bandwidth",
        "Most_Active_Clients_Clients_by_Hits",
        "ALL_Most_Active_Clients_Clients_by_Bandwidth",
        "ALL_Most_Active_Clients_Clients_by_Hits",
    ],
    "pop_domain": [
        "Most_Popular_Domains_Hits",
        "Most_Popular_Domains_Bytes",
        "ALL_Most_Popular_Domains_Bytes",
        "ALL_Most_Popular_Domains_Hits",
    ],
    "top_cli_user": [
        "Top_Clients_Users__Sent_and_Received__by_Bandwidth",
        "Top_Clients_Users_(Sent_and_Received)_by_Bandwidth",
    ],
    "top_cli_host": [
        "Top_Clients_Hosts__Sent_and_Received__by_Bandwidth",
        "Top_Clients_Hosts_(Sent_and_Received)_by_Bandwidth",
    ],
    "top_cli_hits": "Top_Clients_Hosts_by_Hits",
    "GAV": ["Virus_(GAV)_Activity_Trend", "Virus_Activity_Trend"],
    "AD": [
        "AD_Active_Users",
        "AD_Active_Computers",
        "AD_Admin_Users",
        "AD_SSLVPN_Users",
    ],
    "backup": "Backup.pdf",
    "critical": "Critical Events.pdf",
    "dist_group": ["Distribution Group Members", "DistributionGroupMembers"],
    "forward": ["Forwarding Users", "forwardingusers"],
    "MAV": "MAV.pdf",
    "365": ["o365Mail", "o365Apps"],
    "patch": "Patch",
    "server": "ServerWeekly",
    "take_control": ["Take Control", "TakeControl"],
    "web_protect": ["Web Protection", "WebProtection"],
    "TDR": ["WatchGuardTDR", "TDR_Executive Summary", "Executive_Summary_Report"],
    "sslvpn_a_d": ["SSLVPN Authentication Allowed", "SSLVPN Authentication Denied"],
    "proxy": [
        "Proxy_Traffic_Destination_by_Bandwidth",
        "Proxy_Traffic_Destination_by_Hits",
    ],
    "cover_sheet": "CoverSheet.pdf",
    "APT": ["APT__Activity_Trend", "APT)_Activity_Trend"],
    "RED": [
        "Reputation_Enabled_Defense_Action",
        "Reputation_Enabled_Defense_Activity_Trend",
    ],
    "DLP": [
        "DLP__Activity_Trend",
        "DLP__Recipient_Destination",
        "DLP__Rules",
        "DLP__Sender_Source",
        "DLP Details",
    ],
    "SO Review": ["Weekly Service"],
}

# Order of documents for merging final reports
intermax = [
    f_key["cover_sheet"],
    f_key["backup"],
    f_key["365"][0],
    f_key["365"][1],
    f_key["take_control"][0],
    f_key["take_control"][1],
    f_key["critical"],
    f_key["server"],
    f_key["patch"],
    f_key["MAV"],
    f_key["web_protect"][0],
    f_key["web_protect"][1],
    f_key["top_cli_host"][1],
    f_key["top_cli_host"][0],
    f_key["top_cli_user"][1],
    f_key["top_cli_user"][0],
    f_key["app_use_sum"],
    f_key["app_use_bw"][0],
    f_key["app_use_bw"][1],
    f_key["block_sites"][0],
    f_key["block_sites"][1],
    f_key["botnets"][0],
    f_key["botnet_detect"][0],
    f_key["botnet_detect"][1],
    f_key["botnets"][2],
    f_key["IPS"][0],
    f_key["IPS"][1],
    f_key["IPS"][2],
    f_key["IPS"][3],
    f_key["IPS"][4],
    f_key["IPS"][5],
    f_key["IPS"][6],
    f_key["GAV"][0],
    f_key["GAV"][1],
    f_key["AD"][2],
    f_key["AD"][0],
    f_key["AD"][1],
    f_key["AD"][3],
]
knudtsen = [
    f_key["cover_sheet"],
    f_key["backup"],
    f_key["365"][0],
    f_key["365"][1],
    f_key["take_control"][0],
    f_key["take_control"][1],
    f_key["critical"],
    f_key["server"],
    f_key["patch"],
    f_key["MAV"],
    f_key["web_protect"][0],
    f_key["web_protect"][1],
    f_key["top_cli_host"][1],
    f_key["top_cli_host"][0],
    f_key["top_cli_user"][1],
    f_key["top_cli_user"][0],
    f_key["app_use_bw"][0],
    f_key["app_use_bw"][1],
    f_key["active_client"][1],
    f_key["pop_domain"][1],
    f_key["pop_domain"][0],
    f_key["block_sites"][0],
    f_key["block_sites"][1],
    f_key["botnets"][0],
    f_key["botnet_detect"][0],
    f_key["botnet_detect"][1],
    f_key["botnets"][2],
    f_key["AD"][3],
    f_key["AD"][2],
    f_key["AD"][0],
    f_key["AD"][1],
]
bankcda = [
    # f_key["cover_sheet"],
    f_key["backup"],
    f_key["365"][0],
    f_key["365"][1],
    f_key["take_control"][0],
    f_key["take_control"][1],
    f_key["critical"],
    f_key["server"],
    f_key["patch"],
    f_key["MAV"],
    f_key["web_protect"][0],
    f_key["web_protect"][1],
    f_key["active_client"][2],
    f_key["active_client"][3],
    f_key["pop_domain"][2],
    f_key["pop_domain"][3],
    f_key["app_use_bw"][2],
    f_key["block_sites"][2],
    f_key["IPS"][8],
    f_key["TDR"][0],
    f_key["SO Review"][0],
]
honi = [
    f_key["cover_sheet"],
    f_key["backup"],
    f_key["365"][0],
    f_key["365"][1],
    f_key["take_control"][0],
    f_key["take_control"][1],
    f_key["critical"],
    f_key["server"],
    f_key["patch"],
    f_key["MAV"],
    f_key["web_protect"][0],
    f_key["web_protect"][1],
    f_key["top_cli_host"][1],
    f_key["top_cli_host"][0],
    f_key["top_cli_hits"],
    f_key["pop_domain"][1],
    f_key["pop_domain"][0],
    f_key["app_use_bw"][0],
    f_key["app_use_bw"][1],
    f_key["active_client"][0],
    f_key["active_client"][1],
    f_key["block_sites"][0],
    f_key["block_sites"][1],
    f_key["IPS"][0],
    f_key["IPS"][1],
    f_key["IPS"][2],
    f_key["IPS"][3],
    f_key["IPS"][4],
    f_key["IPS"][5],
    f_key["IPS"][6],
    f_key["IPS"][7],
    f_key["GAV"][0],
    f_key["GAV"][1],
    f_key["botnets"][0],
    f_key["botnet_detect"][0],
    f_key["botnet_detect"][1],
    f_key["botnets"][2],
    f_key["sslvpn_a_d"][0],
    f_key["sslvpn_a_d"][1],
    f_key["TDR"][0],
    f_key["TDR"][1],
    f_key["AD"][0],
    f_key["AD"][1],
    f_key["AD"][2],
    f_key["AD"][3],
]
mmco = [
    f_key["cover_sheet"],
    f_key["backup"],
    f_key["365"][0],
    f_key["365"][1],
    f_key["take_control"][0],
    f_key["take_control"][1],
    f_key["critical"],
    f_key["server"],
    f_key["patch"],
    f_key["MAV"],
    f_key["web_protect"][0],
    f_key["web_protect"][1],
    f_key["top_cli_host"][1],
    f_key["top_cli_host"][0],
    f_key["top_cli_user"][1],
    f_key["top_cli_user"][0],
    f_key["app_use_bw"][0],
    f_key["app_use_bw"][1],
    f_key["active_client"][1],
    f_key["pop_domain"][1],
    f_key["pop_domain"][0],
    f_key["block_sites"][0],
    f_key["block_sites"][1],
    f_key["botnets"][0],
    f_key["botnet_detect"][0],
    f_key["botnet_detect"][1],
    f_key["botnets"][2],
    f_key["TDR"][2],
    f_key["sslvpn_a_d"][0],
    f_key["sslvpn_a_d"][1],
    f_key["AD"][3],
    f_key["AD"][2],
    f_key["AD"][0],
    f_key["AD"][1],
]
integrated = [
    f_key["cover_sheet"],
    f_key["backup"],
    f_key["365"][0],
    f_key["365"][1],
    f_key["take_control"][0],
    f_key["take_control"][1],
    f_key["critical"],
    f_key["server"],
    f_key["patch"],
    f_key["MAV"],
    f_key["web_protect"][0],
    f_key["web_protect"][1],
    f_key["top_cli_host"][1],
    f_key["top_cli_host"][0],
    f_key["top_cli_hits"],
    f_key["app_use_bw"][0],
    f_key["app_use_bw"][1],
    f_key["app_use_hits"][0],
    f_key["app_use_hits"][1],
    f_key["app_use_hits"][2],
    f_key["pop_domain"][1],
    f_key["pop_domain"][0],
    f_key["active_client"][0],
    f_key["active_client"][1],
    f_key["block_sites"][0],
    f_key["IPS"][0],
    f_key["IPS"][1],
    f_key["GAV"][0],
    f_key["GAV"][1],
    f_key["botnets"][0],
    f_key["botnet_detect"][0],
    f_key["botnet_detect"][1],
    f_key["botnets"][2],
    f_key["sslvpn_a_d"][0],
    f_key["sslvpn_a_d"][1],
    f_key["AD"][1],
    f_key["AD"][0],
    f_key["AD"][2],
    f_key["AD"][3],
    f_key["dist_group"][0],
    f_key["dist_group"][1],
    f_key["forward"][0],
    f_key["forward"][1],
]
northcon = [
    f_key["cover_sheet"],
    f_key["backup"],
    f_key["365"][0],
    f_key["365"][1],
    f_key["take_control"][0],
    f_key["take_control"][1],
    f_key["critical"],
    f_key["server"],
    f_key["patch"],
    f_key["MAV"],
    f_key["top_cli_host"][1],
    f_key["top_cli_host"][0],
    f_key["top_cli_hits"],
    f_key["app_cont"],
    f_key["app_use_bw"][0],
    f_key["app_use_bw"][1],
    f_key["block_sites"][0],
    f_key["IPS"][0],
    f_key["IPS"][1],
    f_key["IPS"][2],
    f_key["IPS"][3],
    f_key["IPS"][4],
    f_key["IPS"][5],
    f_key["IPS"][6],
    f_key["GAV"][0],
    f_key["GAV"][1],
    f_key["proxy"][0],
    f_key["proxy"][1],
    f_key["botnets"][0],
    f_key["botnet_detect"][0],
    f_key["botnet_detect"][1],
    f_key["botnets"][2],
    f_key["sslvpn_a_d"][0],
    f_key["sslvpn_a_d"][1],
    f_key["TDR"][0],
    f_key["TDR"][1],
    f_key["AD"][2],
    f_key["AD"][0],
    f_key["AD"][3],
    f_key["AD"][1],
]
bayshore = [
    f_key["cover_sheet"],
    f_key["backup"],
    f_key["365"][0],
    f_key["365"][1],
    f_key["take_control"][0],
    f_key["take_control"][1],
    f_key["critical"],
    f_key["server"],
    f_key["patch"],
    f_key["MAV"],
    f_key["web_protect"][0],
    f_key["web_protect"][1],
    f_key["top_cli_host"][1],
    f_key["top_cli_host"][0],
    f_key["top_cli_user"][1],
    f_key["top_cli_user"][0],
    f_key["active_client"][0],
    f_key["app_use_bw"][0],
    f_key["app_use_bw"][1],
    f_key["pop_domain"][1],
    f_key["pop_domain"][0],
    f_key["block_sites"][0],
    f_key["block_sites"][1],
    f_key["botnets"][0],
    f_key["botnet_detect"][0],
    f_key["botnet_detect"][1],
    f_key["botnets"][2],
    f_key["IPS"][0],
    f_key["IPS"][1],
    f_key["GAV"][0],
    f_key["GAV"][1],
    f_key["sslvpn_a_d"][0],
    f_key["sslvpn_a_d"][1],
    f_key["TDR"][0],
    f_key["TDR"][1],
    f_key["AD"][0],
    f_key["AD"][1],
    f_key["AD"][2],
    f_key["AD"][3],
]
pffm = [
    f_key["cover_sheet"],
    f_key["backup"],
    f_key["365"][0],
    f_key["365"][1],
    f_key["take_control"][0],
    f_key["take_control"][1],
    f_key["critical"],
    f_key["server"],
    f_key["patch"],
    f_key["MAV"],
    f_key["web_protect"][0],
    f_key["web_protect"][1],
    f_key["top_cli_host"][1],
    f_key["top_cli_host"][0],
    f_key["active_client"][0],
    f_key["pop_domain"][1],
    f_key["pop_domain"][0],
    f_key["app_use_bw"][0],
    f_key["app_use_bw"][1],
    f_key["block_sites"][0],
    f_key["block_sites"][1],
    f_key["botnets"][0],
    f_key["botnet_detect"][0],
    f_key["botnets"][2],
    f_key["botnet_detect"][1],
    f_key["IPS"][0],
    f_key["IPS"][1],
    f_key["GAV"][0],
    f_key["GAV"][1],
    f_key["sslvpn_a_d"][0],
    f_key["sslvpn_a_d"][1],
    f_key["AD"][0],
    f_key["AD"][1],
    f_key["AD"][2],
    f_key["AD"][3],
]

# List of clients
client_list = [
    ["Intermax", intermax],
    ["Knudtsen", knudtsen],
    ["BankCDA", bankcda],
    ["HONI", honi],
    ["MMCO", mmco],
    ["Integrated Personnel", integrated],
    ["Northcon", northcon],
    ["Bay Shore", bayshore],
    ["PFFM", pffm],
]

#############
# FUNCTIONS #
#############

# Select the client for the merged report
def select_client():
    i = 1
    for item in client_list:
        print(f"{i}) {item[0]}")  # Print out a list of clients to select from
        i += 1
    client = int(input("Please select the client by number:\n>"))
    return client - 1  # Return the client as an index


# Create a list of all the weekly folder paths
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
        "\\\\FS01\\MSP-SecReview\\weekly"  # Upper level directory for all client files
    )
    dir_list = listdir(dir_path)

    paths = []  # Create a list of all the current week directories
    for item in dir_list:
        if item[0].isdigit():
            paths.append(
                dir_path + "\\" + item + "\\" + str_date_monday
            )  # Add each file to the list as a full filepath
    return paths


# Convert the cover sheet to a pdf
def convert_docx(folder_path):
    wdFormatPDF = 17
    file_in = folder_path + "\\" + "_CoverSheet.docx"  # Path to the input docx file
    file_out = folder_path + "\\" + "CoverSheet.pdf"  # Path to the output pdf file
    inputFile = os.path.abspath(file_in)
    outputFile = os.path.abspath(file_out)
    word = client.Dispatch("Word.Application")
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


# Create a list of all the files in the specific client weekly directory
def dir_list(client_selection):
    all_files = []
    folders = all_dir_paths()
    cli_dir_path = folders[
        client_selection
    ]  # Concatenate a path to the specific weekly folder for the selected client
    print(cli_dir_path)
    print("Converting CoverSheet.docx")
    convert_docx(cli_dir_path)  # Convert the
    file_list = listdir(
        cli_dir_path
    )  # Create a list of all the files in that directory
    for item in file_list:
        path = cli_dir_path + "\\" + item
        all_files.append(path)  # Append all file paths to the files list
    return all_files, cli_dir_path


# Creates a list of the files to be used in the merged report. Find files from dir_list based on f_key keywords
def file_list(client_keys, pdfs):
    file_paths = []
    for search_word in client_keys:
        for item in pdfs:  # For each item in the pdfs file list,
            if (
                search_word in item
            ):  # Check to see if there is a match with the given entry in files
                file_paths.append(item)  # Add the filepath to the items to merge
    return file_paths


# Creates a string of the date range to be used for the report filename
def dates_for_report():
    day_of_week = date.today().weekday()  # Returns the day of the week as an integer
    day_delta_start = day_of_week + 7  # Find the starting day of the report as a delta
    day_delta_end = day_of_week + 1  # Find the ending day of the report as a delta
    starting = date.today() - timedelta(days=day_delta_start)  # Grab full date start
    ending = date.today() - timedelta(days=day_delta_end)  # Grab full date end
    start_date = starting.strftime("%Y-%m-%d")  # Extract only the starting yyyy-mm-dd
    end_date = ending.strftime("%m-%d")  # Extract only the ending mm-dd
    date_range = start_date + "-" + end_date  # Concatenate the start and end
    return date_range


# Merge all of the present files
def merge_pdfs(
    pdfs_list, client_selection
):  # Arguments: 1) Complete list of pdfs for specific client; 2) Path to that client's weekly folder
    f = open("Merged_files.txt", "a")
    date_range = dates_for_report()  # Get the date for the security report
    merger = PdfFileMerger(strict=False)  # Ignore bad white space
    result = date_range + " Weekly Security Report.pdf"  # Output file path

    for pdf in pdfs_list:
        # print(f"Adding {pdf} to the merged pdf...")
        f.write(f"{pdf}\n")
        merger.append(pdf)  # Add all files to the merger
    merger.write(result)  # Create the merged pdf
    merger.close()
    print("Finished")
    f.close()
    folders = all_dir_paths()
    dest_dir = folders[
        client_selection
    ]  # Concatenate a path to the specific weekly folder for the selected client
    dest = dest_dir + "\\" + result
    shutil.move(result, dest)


### MAIN ###
client_selection = select_client()  # Select the client directory
pdfs = dir_list(
    client_selection
)  # Create a list of all the files in that directory and convert Cover Sheet
keys = client_list[client_selection][1]  # Grab keys based off of client selection
final_files = file_list(keys, pdfs[0])  # Create list of the files to merge
merge_pdfs(final_files, client_selection)  # Merge files
