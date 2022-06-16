import PyPDF2
from PyPDF2 import PdfFileMerger
import os
from win32com import client
from datetime import date, timedelta


# Dictionary of keywords for the pdf files
f_key = {"app_use_sum": "Application_Control_Application_Usage_Summary",
"app_cont": "Application_Control_Blocked_Application_Summary",
"app_use_bw": ["Application_Usage_Applications_(Host)_by_Bandwidth", "Application_Usage_Applications__Host__by_Bandwidth"],
"app_use_hits": ["Application_Usage_Applications_(Host)_by_Hits", "Application_Usage_Applications__Host__by_Hits", "Application_Usage_Top_Hosts_by_Hits"],
"block_sites": ["Blocked_Websites_Category", "Blocked_Websites_Client"],
"botnets": ["Botnet_Detection_Activity_Trend", "BotNet Detection Details", "Blocked_Botnet_Sites"],
"botnet_detect": ["Botnet_Detection_Botnet_Detection_by_Client", "Botnet_Detection_Botnet_Detection_by_Destination"],
"IPS": ["Intrusions_(IPS)_Activity_Trend", "Intrusions__IPS__Activity_Trend", "Intrusions__IPS__Source", "Intrusions__IPS__Signatures", "Intrusions__IPS__Protocol", "Intrusions__IPS__Threat_Level", "Intrusions Prevention Details"],
"active_client": ["Most_Active_Clients_Clients_by_Bandwidth", "Most_Active_Clients_Clients_by_Hits"],
"pop_domain": ["Most_Popular_Domains_Hits", "Most_Popular_Domains_Bytes"],
"top_cli_user": ["Top_Clients_Users__Sent_and_Received__by_Bandwidth", "Top_Clients_Users_(Sent_and_Received)_by_Bandwidth"],
"top_cli_host": ["Top_Clients_Hosts__Sent_and_Received__by_Bandwidth", "Top_Clients_Hosts_(Sent_and_Received)_by_Bandwidth"],
"top_cli_hits": "Top_Clients_Hosts_by_Hits",
"top_cli_user": "Top_Clients_Users_by_Hits",
"GAV": ["Virus_(GAV)_Activity_Trend", "Virus_Activity_Trend"],
"AD": ["AD_Active_Users", "AD_Active_Computers", "AD_Admin_Users", "AD_SSLVPN_Users"],
"backup": "Backup.pdf",
"critical": "Critical Events.pdf",
"dist_group": ["Distribution Group Members", "DistributionGroupMembers"],
"forward": ["Forwarding Users", "forwardingusers"],
"MAV": "MAV.pdf",
"365": ["o365Mail", "o365Apps", "Office365Usage"],
"patch": "Patch",
"server": "ServerWeekly",
"take_control": ["Take Control", "TakeControl"],
"web_protect": ["Web Protection","WebProtection"],
"TDR": "WatchGuardTDR",
"sslvpn_a_d": ["SSLVPN Authentication Allowed", "SSLVPN Authentication Denied"],
"proxy": ["Proxy_Traffic_Destination_by_Bandwidth", "Proxy_Traffic_Destination_by_Hits"],
"cover_sheet": "CoverSheet.pdf",
"APT": ["APT__Activity_Trend", "APT)_Activity_Trend"],
"RED": ["Reputation_Enabled_Defense_Action", "Reputation_Enabled_Defense_Activity_Trend"],
"DLP": ["DLP__Activity_Trend", "DLP__Recipient_Destination", "DLP__Rules", "DLP__Sender_Source", "DLP Details"]
}

intermax = [f_key["cover_sheet"], f_key["backup"], f_key["365"][0], f_key["365"][1], f_key["take_control"][0], f_key["take_control"][1], f_key["critical"], f_key["server"], f_key["patch"], f_key["MAV"], f_key["web_protect"][0], f_key["web_protect"][1], f_key["top_cli_host"][1], f_key["top_cli_host"][0], f_key["top_cli_user"][1], f_key["top_cli_user"][0], f_key["app_use_sum"], f_key["app_use_bw"][0], f_key["app_use_bw"][1], f_key["block_sites"][0], f_key["block_sites"][1],  f_key["botnets"][0], f_key["botnet_detect"][0], f_key["botnet_detect"][1], f_key["botnets"][2], f_key["IPS"][0], f_key["IPS"][1], f_key["IPS"][2], f_key["IPS"][3], f_key["IPS"][4], f_key["IPS"][5], f_key["IPS"][6], f_key["GAV"][0], f_key["GAV"][1], f_key["AD"][2], f_key["AD"][0], f_key["AD"][1], f_key["AD"][3]]
knudtsen = [f_key["cover_sheet"], f_key["backup"], f_key["365"][0], f_key["365"][1], f_key["take_control"][0], f_key["take_control"][1], f_key["critical"], f_key["server"], f_key["patch"], f_key["MAV"], f_key["web_protect"][0], f_key["web_protect"][1], f_key["top_cli_host"][1], f_key["top_cli_host"][0], f_key["top_cli_user"][1], f_key["top_cli_user"][0], f_key["app_use_bw"][0], f_key["app_use_bw"][1], f_key["active_client"][1], f_key["pop_domain"][1], f_key["pop_domain"][0], f_key["block_sites"][0], f_key["block_sites"][1], f_key["botnets"][0], f_key["botnet_detect"][0], f_key["botnet_detect"][1], f_key["botnets"][2], f_key["AD"][3], f_key["AD"][2], f_key["AD"][0], f_key["AD"][1]]
bankcda = [f_key["cover_sheet"], f_key["backup"], f_key["365"][0], f_key["365"][1], f_key["take_control"][0], f_key["take_control"][1], f_key["critical"], f_key["server"], f_key["patch"], f_key["MAV"], f_key["web_protect"][0], f_key["web_protect"][1], f_key["top_cli_host"][1], f_key["top_cli_host"][0], f_key["top_cli_user"][1], f_key["top_cli_user"][0], f_key["app_use_bw"][0], f_key["app_use_bw"][1], f_key["active_client"][1], f_key["pop_domain"][1], f_key["pop_domain"][0], f_key["block_sites"][0], f_key["block_sites"][1], f_key["botnets"][0], f_key["botnet_detect"][0], f_key["botnet_detect"][1], f_key["botnets"][2], f_key["AD"][3], f_key["AD"][2], f_key["AD"][0], f_key["AD"][1]]
honi = [f_key["cover_sheet"], f_key["backup"], f_key["365"][0], f_key["365"][1], f_key["take_control"][0], f_key["take_control"][1], f_key["critical"], f_key["server"], f_key["patch"], f_key["MAV"], f_key["web_protect"][0], f_key["web_protect"][1], f_key["top_cli_host"][1], f_key["top_cli_host"][0], f_key["top_cli_user"], f_key["pop_domain"][1], f_key["pop_domain"][0], f_key["app_use_bw"][0], f_key["app_use_bw"][1], f_key["active_client"][0], f_key["active_client"][1], f_key["block_sites"][0], f_key["block_sites"][1], f_key["IPS"][0], f_key["IPS"][1], f_key["IPS"][2], f_key["IPS"][3], f_key["IPS"][4], f_key["IPS"][5], f_key["IPS"][6], f_key["GAV"][0], f_key["GAV"][1], f_key["botnets"][0], f_key["botnet_detect"][0], f_key["botnet_detect"][1], f_key["botnets"][2], f_key["sslvpn_a_d"][0], f_key["sslvpn_a_d"][1], f_key["TDR"], f_key["AD"][0], f_key["AD"][1], f_key["AD"][2], f_key["AD"][3]]
mmco = [f_key["cover_sheet"], f_key["backup"], f_key["365"][0], f_key["365"][1], f_key["take_control"][0], f_key["take_control"][1], f_key["critical"], f_key["server"], f_key["patch"], f_key["MAV"], f_key["web_protect"][0], f_key["web_protect"][1], f_key["top_cli_host"][1], f_key["top_cli_host"][0], f_key["top_cli_user"][1], f_key["top_cli_user"][0], f_key["app_use_bw"][0], f_key["app_use_bw"][1], f_key["active_client"][1], f_key["pop_domain"][1], f_key["pop_domain"][0], f_key["block_sites"][0], f_key["block_sites"][1], f_key["botnets"][0], f_key["botnet_detect"][0], f_key["botnet_detect"][1], f_key["botnets"][2], f_key["AD"][3], f_key["AD"][2], f_key["AD"][0], f_key["AD"][1]]
integrated = [f_key["cover_sheet"], f_key["backup"], f_key["365"][0], f_key["365"][1], f_key["take_control"][0], f_key["take_control"][1], f_key["critical"], f_key["server"], f_key["patch"], f_key["MAV"], f_key["web_protect"][0], f_key["web_protect"][1], f_key["top_cli_host"][1], f_key["top_cli_host"][0], f_key["top_cli_hits"], f_key["app_use_bw"][0], f_key["app_use_bw"][1], f_key["app_use_hits"][0], f_key["app_use_hits"][1], f_key["app_use_hits"][2], f_key["pop_domain"][1], f_key["pop_domain"][0],f_key["active_client"][0], f_key["active_client"][1], f_key["block_sites"][0], f_key["IPS"][0], f_key["IPS"][1], f_key["GAV"][0], f_key["GAV"][1], f_key["botnets"][0], f_key["botnet_detect"][0], f_key["botnet_detect"][1], f_key["botnets"][2], f_key["AD"][1], f_key["AD"][0], f_key["AD"][2], f_key["AD"][3], f_key["dist_group"][0], f_key["dist_group"][1], f_key["forward"][0], f_key["forward"][1]]
northcon = [f_key["cover_sheet"], f_key["backup"], f_key["365"][0], f_key["365"][1], f_key["take_control"][0], f_key["take_control"][1], f_key["critical"], f_key["server"], f_key["patch"], f_key["MAV"], f_key["top_cli_host"][1], f_key["top_cli_host"][0], f_key["top_cli_hits"], f_key["app_cont"], f_key["app_use_bw"][0], f_key["app_use_bw"][1], f_key["block_sites"][0], f_key["IPS"][0], f_key["IPS"][1], f_key["IPS"][2], f_key["IPS"][3], f_key["IPS"][4], f_key["IPS"][5], f_key["IPS"][6], f_key["GAV"][0], f_key["GAV"][1], f_key["RED"][1], f_key["RED"][0], f_key["APT"][0], f_key["APT"][1], f_key["proxy"][0], f_key["proxy"][1], f_key["botnets"][0], f_key["botnet_detect"][0], f_key["botnet_detect"][1], f_key["botnets"][2], f_key["sslvpn_a_d"][0], f_key["sslvpn_a_d"][1], f_key["TDR"], f_key["AD"][2], f_key["AD"][0], f_key["AD"][3], f_key["AD"][1]]
bayshore = [f_key["cover_sheet"], f_key["backup"], f_key["365"][0], f_key["365"][1], f_key["take_control"][0], f_key["take_control"][1], f_key["critical"], f_key["server"], f_key["patch"], f_key["MAV"], f_key["web_protect"][0], f_key["web_protect"][1], f_key["top_cli_host"][1], f_key["top_cli_host"][0], f_key["top_cli_user"][1], f_key["top_cli_user"][0], f_key["active_client"][0], f_key["app_use_bw"][0], f_key["app_use_bw"][1], f_key["pop_domain"][1], f_key["pop_domain"][0], f_key["block_sites"][0], f_key["block_sites"][1], f_key["botnets"][0], f_key["botnet_detect"][0], f_key["botnet_detect"][1], f_key["botnets"][2], f_key["IPS"][0], f_key["IPS"][1], f_key["GAV"][0], f_key["GAV"][1], f_key["RED"][0], f_key["RED"][1], f_key["APT"][0], f_key["APT"][1], f_key["sslvpn_a_d"][0], f_key["sslvpn_a_d"][1], f_key["TDR"], f_key["AD"][0], f_key["AD"][1], f_key["AD"][2], f_key["AD"][3]]
mpms = [f_key["cover_sheet"], f_key["backup"], f_key["365"][0], f_key["365"][1], f_key["take_control"][0], f_key["take_control"][1], f_key["critical"], f_key["server"], f_key["patch"], f_key["MAV"], f_key["web_protect"][0], f_key["web_protect"][1], f_key["top_cli_host"][1], f_key["top_cli_host"][0], f_key["top_cli_user"][1], f_key["top_cli_user"][0], f_key["pop_domain"][0], f_key["pop_domain"][1], f_key["active_client"][0], f_key["active_client"][1], f_key["app_use_bw"][0], f_key["app_use_bw"][1], f_key["block_sites"][0], f_key["block_sites"][1], f_key["botnets"][0], f_key["botnet_detect"][0], f_key["botnets"][2], f_key["botnet_detect"][1], f_key["GAV"][0], f_key["GAV"][1], f_key["IPS"][0], f_key["IPS"][1], f_key["AD"][1], f_key["AD"][0], f_key["AD"][2], f_key["AD"][3]]
pffm = [f_key["cover_sheet"], f_key["backup"], f_key["365"][0], f_key["365"][1], f_key["take_control"][0], f_key["take_control"][1], f_key["critical"], f_key["server"], f_key["patch"], f_key["MAV"], f_key["web_protect"][0], f_key["web_protect"][1], f_key["top_cli_host"][1], f_key["top_cli_host"][0], f_key["active_client"][0], f_key["pop_domain"][1], f_key["pop_domain"][0], f_key["app_use_bw"][0], f_key["app_use_bw"][1], f_key["block_sites"][0], f_key["block_sites"][1], f_key["botnets"][0], f_key["botnet_detect"][0], f_key["botnets"][2], f_key["botnet_detect"][1], f_key["IPS"][0], f_key["IPS"][1], f_key["GAV"][0], f_key["GAV"][1], f_key["RED"][1], f_key["RED"][0], f_key["DLP"][0], f_key["DLP"][2], f_key["DLP"][3], f_key["DLP"][1], f_key["DLP"][4], f_key["AD"][0], f_key["AD"][1], f_key["AD"][2], f_key["AD"][3]]

# client_list = {"Intermax": intermax, "Knudtsen": knudtsen, "BankCDA": bankcda, "HONI": honi, "MMCO": mmco, "Integrated Personnel": integrated, "Northcon": northcon, "Bay Shore": bayshore, "MPMS": mpms, "PFFM": pffm}

client_list = [["Intermax", intermax], ["Knudtsen", knudtsen], ["BankCDA", bankcda], ["HONI", honi], ["MMCO", mmco], ["Integrated Personnel", integrated], ["Northcon", northcon], ["Bay Shore", bayshore], ["MPMS", mpms], ["PFFM", pffm]]

def select_client(client_list):
    i = 1
    for client in client_list:
        print(f"{i}) {client[0]}")
        i += 1
    client_choice = int(input("Please select a client by number:\n>")) - 1
    return client_list[client_choice][1]
        
# Rotate the MAV pdf, if needed
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
def file_list(client_keys, pdfs):
    print("Compiling a list of all the pertinent files which are present...")
    file_paths = []
    for search_word in client_keys:
        for item in pdfs:                                                       # For each item in the pdfs file list,
            if search_word in item:                                             # Check to see if there is a match with the given entry in files
                file_paths.append(item)                                         # Add the filepath to the items to merge
    return file_paths

def convert_docx():
    wdFormatPDF = 17
    environment = os.getcwd() + "\\"                                            # Get local directory path
    file_in = environment + "_CoverSheet.docx"
    file_out = environment + "CoverSheet.pdf"
    inputFile = os.path.abspath(file_in)
    outputFile = os.path.abspath(file_out)
    word = client.Dispatch('Word.Application')
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

def dates_for_report():
    # Returns the day of the week as an integer
    day_of_week = date.today().weekday()
    # Add the day to a full week
    day_delta_start = day_of_week + 7
    day_delta_end = day_of_week + 1
    starting = date.today() - timedelta(days=day_delta_start)
    ending = date.today() - timedelta(days=day_delta_end)
    start_date = starting.strftime("%Y-%m-%d")
    end_date = ending.strftime("%m-%d")
    date_range = start_date + "-" + end_date
    return date_range

# Merge all of the present files
def merge_pdfs(pdfs_list):
    date_range = dates_for_report()
    merger = PdfFileMerger(strict=False)                                        # Ignore bad white space
    result = os.getcwd() + "\\" + date_range + " Weekly Security Report.pdf"   # Output file path
    for pdf in pdfs_list:
        print(f"Adding {pdf} to the merged pdf...")
        merger.append(pdf)                                                      # Add all files to the merger
    print("Saving the merged pdf...")
    merger.write(result)                                                        # Create the merged pdf
    merger.close()

### MAIN ###
convert_docx()
pdfs = pdf_paths()
keys = select_client(client_list)
final_files = file_list(keys, pdfs)
merge_pdfs(final_files)
