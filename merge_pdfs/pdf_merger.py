# from threading import local
import PyPDF2
from PyPDF2 import PdfFileMerger
import os

files_general = [
"Backup.pdf",
"o365Mail",
"o365Apps",
"Take Control",
"TakeControl",
"Critical Events",
"ServerWeekly",
"PatchManagement",
"mav_rotated.pdf",
"Web Protection",
"Top_Clients_Hosts__Sent_and_Received__by_Bandwidth",
"Top_Clients_Users__Sent_and_Received__by_Bandwidth",
"Usage_Applications_(Host)_by_Bandwidth",
"Active_Clients_Clients_by_Bandwidth",
"Popular_Domains_Bytes",
"Popular_Domains_Hits",
"Blocked_Websites_Category",
"Blocked_Websites_Client",
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
# PFFM
files = [
"Backup.pdf",
"o365Mail",
"o365Apps",
"Take Control",
"TakeControl",
"Critical Events",
"ServerWeekly",
"PatchManagement",
"mav_rotated.pdf",
"Web Protection",
"Top_Clients_Hosts__Sent_and_Received__by_Bandwidth",
"Top_Clients_Users__Sent_and_Received__by_Bandwidth",
"Active_Clients_Clients_by_Bandwidth",
"Usage_Applications_(Host)_by_Bandwidth",
"Popular_Domains_Bytes",
"Popular_Domains_Hits",
"Blocked_Websites_Category",
"Blocked_Websites_Client",
"Botnet_Detection_Activity_Trend",
"Botnet_Detection_by_Client",
"Botnet_Detection_by_Destination",
"Blocked_Botnet_Sites",
"BotNet Detection Details",
"AD_SSLVPN_Users",
"AD_Admin_Users",
"AD_Active_Users",
"AD_Active_Computers"
]

# Rotate the MAV pdf
def rotate():
    print("Rotating the MAV pdf...")
    pdf_in = open('MAV.pdf', 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_in,strict=False)
    pdf_writer = PyPDF2.PdfFileWriter()

    for pagenum in range(pdf_reader.numPages):
        page = pdf_reader.getPage(pagenum)
        page.rotateClockwise(90)
        pdf_writer.addPage(page)

    pdf_out = open('mav_rotated.pdf', 'wb')
    pdf_writer.write(pdf_out)
    pdf_out.close()
    pdf_in.close()

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
    print("Compiling a list of all the pertinent files which are present...")
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
        print(f"Adding {pdf} to the merged pdf...")
        merger.append(pdf)                                                      # Add all files to the merger
    print("Saving the merged pdf...")
    merger.write(result)                                                        # Create the merged pdf
    merger.close()
### MAIN ###
rotate()
pdfs = pdf_paths()
final_files = file_list(files, pdfs)
merge_pdfs(final_files)
