from threading import local
from PyPDF2 import PdfFileMerger
import os


files = [
"Backup.pdf",
"o365Mail",
"o365Apps",
"Take Control",
"Critical Events",
"ServerWeekly",
"PatchManagement",
"MAV.pdf",
"Web Protection",
"Top_Clients_Hosts",
"Clients_by_Bandwidth",
"Domains_Hits",
"Application_Usage_Application",
"Blocked_Websites_Category",
"Botnet_Detection_Activity_Trend",
"Botnet_Detection_by_Client",
"Botnet_Detection_by_Destination",
"Blocked_Botnet_Sites",
"BotNet Detection Details",
"Intrusions_(IPS)",
"Virus_(GAV)",
"AD_Active_Users",
"AD_Active_Computers",
"AD_Admin_Users",
"AD_SSLVPN_Users"
]

# Create list of file paths for the pdfs
def pdf_paths():                                                                # Create list of file paths for the pdfs
    pdfs = os.listdir()                                                         # Get list of all files in the directory
    environment = os.getcwd() + "\\"                                            # Get local directory path
    i = 0
    for item in pdfs:
        pdfs[i] = environment + item                                            # Change the list entries to their full file path
        i += 1
    return pdfs

# Create the final list of filtered file paths: Include full path if there is a match present
def file_list(files, pdfs):
    file_paths = []
    for search_word in files:
        for item in pdfs:                                                       # For each item in the pdfs file list,
            if search_word in item:                                             # Check to see if there is a match with the given entry in files
                file_paths.append(item)                                         # Add the filepath to the items to merge
    return file_paths

# Merge all of the present files
def merge_pdfs(pdfs_list):
    merger = PdfFileMerger(strict=False)                                        # Ignore bad white space
    result = os.getcwd() + "\\result.pdf"                                       # Output file path
    for pdf in pdfs_list:
        merger.append(pdf)                                                      # Add all files to the merger
    merger.write(result)                                                        # Create the merged pdf
    merger.close()

### MAIN ###
pdfs = pdf_paths()
final_files = file_list(files, pdfs)
merge_pdfs(final_files)
