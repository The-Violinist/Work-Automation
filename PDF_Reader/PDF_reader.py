from datetime import date, timedelta
from os import listdir, remove
from PyPDF2 import PdfFileReader
import os
import re
### VARIABLES ###
# Dictionary of keywords to search for files in the weekly directories
f_type = {
"app_use_bw": ["Application_Usage_Applications_(Host)_by_Bandwidth", "Application_Usage_Applications__Host__by_Bandwidth"],
"app_use_hits": ["Application_Usage_Applications_(Host)_by_Hits", "Application_Usage_Applications__Host__by_Hits", "Application_Usage_Top_Hosts_by_Hits"],
"block_sites_cat": ["Blocked_Websites_Category"],
"block_sites_cli": ["Blocked_Websites_Client"],
"botnet_trend": ["Botnet_Detection_Activity_Trend"],
"botnet_details": ["BotNet Detection Details"],
"botnet_block": ["Blocked_Botnet_Sites"],
"botnet_detect_cli": ["Botnet_Detection_Botnet_Detection_by_Client"],
"botnet_detect_dest": ["Botnet_Detection_Botnet_Detection_by_Destination"],
"active_client_bw": ["Most_Active_Clients_Clients_by_Bandwidth"],
"active_client_hit": ["Most_Active_Clients_Clients_by_Hits"],
"pop_domain_bytes": ["Most_Popular_Domains_Bytes"],
"pop_domain_hits": ["Most_Popular_Domains_Hits"],
"top_cli_user": ["Top_Clients_Users__Sent_and_Received__by_Bandwidth", "Top_Clients_Users_(Sent_and_Received)_by_Bandwidth"],
"top_cli_host": ["Top_Clients_Hosts__Sent_and_Received__by_Bandwidth", "Top_Clients_Hosts_(Sent_and_Received)_by_Bandwidth"],
"top_cli_hits": ["Top_Clients_Hosts_by_Hits"],
"IPS": ["Intrusions_(IPS)_Activity_Trend", "Intrusions__IPS__Activity_Trend"],
"GAV": ["Virus_(GAV)_Activity_Trend", "Virus_Activity_Trend"],
#"MAV": ["MAV.pdf"],
#"patch": ["Patch"],
#"take_control": ["Take Control", "TakeControl"],
#"web_protect": ["Web Protection","WebProtection"],
"proxy": ["Proxy_Traffic_Destination_by_Bandwidth", "Proxy_Traffic_Destination_by_Hits"]
}

clients = ["Intermax","Knudtsen","BankCDA","HONI","MMCO","Integrated Personnel","Northcon","Bay Shore","MPMS","PFFM"]

### FUNCTIONS ###

#############################
# Gather files to work with #
#############################

# Select which client
def select_client():
    i = 1
    for item in clients:
        print(f"{i}) {item}")                                                   # Print out a list of clients to select from
        i += 1
    client = int(input("Please select the client by number:\n>"))
    return client - 1                                                           # Return the client as an index

# Create a list of all the weekly folder paths
def all_dir_paths():
    # Get the subdirectory name based on date
    day_of_week = date.today().weekday()                                        # Get today's day of the week as an index
    date_monday = date.today() - timedelta(days=day_of_week)                    # Calculate the date of the Monday of this week
    str_date_monday = date_monday.strftime("%Y-%m-%d")                          # Convert the complete date of Monday to a string
    # Parent for all client directories
    # dir_path = '\\\\FS01\\MSP-SecReview\\weekly'                                # Upper level directory for all client files

    dir_path = 'C:\\Users\\darmstrong\\Desktop\\script_test'                    # Test directory

    dir_list = listdir(dir_path)
    # Create a list of all the current week directories
    paths = []
    for item in dir_list:
        if item[0].isdigit():
            paths.append(dir_path + "\\" + item + "\\" + str_date_monday)
    return paths

# Create a list of all the files in the specific client weekly directory
def dir_list():
    all_files = []
    client = select_client()
    folders = all_dir_paths()
    cli_dir_path = folders[client]                                              # Concatenate a path to the specific weekly folder for that client
    file_list = listdir(cli_dir_path)                                           # Create a list of all the files in that directory
    for item in file_list:
        path = cli_dir_path + "\\" + item
        all_files.append(path)                                                  # Append all file paths to the files list
    return all_files

def paths_in_dict(cli_files):
    for k in f_type:                                                            # Each key in the search word dictionary
        for value in f_type[k]:                                                 # Each value for each key
            for item in cli_files:                                              # Each file in the directory
                if value in item:                                               # If the search word is in the file name, change the value in the dictionary to the filepath
                    f_type[k] = item
######################################################################################

########################
# Extract Data to TXTs #
########################

# Extract data from PDFs
def extract_data(in_file, out_file):
    temp = open(in_file, 'rb')
    PDF_read = PdfFileReader(temp)
    num_pages = PDF_read.getNumPages()

    f = open(out_file, "a")

    # Extract text and write to text file
    for page in range(num_pages):
        numbered_page = PDF_read.getPage(page)
        page_text = numbered_page.extractText()
        f.write(page_text)
    f.close()

    # Remove empty lines
    with open(out_file, "r") as filehandle:
        lines = filehandle.readlines()
    with open(out_file, 'w') as filehandle:
        lines = filter(lambda x: x.strip(), lines)
        filehandle.writelines(lines)
    filehandle.close()

# Create all txt files
def create_temps():
    temps = []
    for k in f_type:
        in_file = f_type[k]                                                 # Input file is the filepath which was added to the dictionary
        out_file = f"{k}.txt"                                               # Output file is the dictionary key with txt file extension    
        try:
            extract_data(in_file, out_file)                                 # If a filepath exists as a value in the dictionary, extract all text to a temp txt file
            temps.append(out_file)
        except:
            pass
    return temps
#######################################################################################

################
# Analyze Data #
################

def IPS(temp_file="IPS.txt"):
    final_data = []
    # Read data into variable
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        int_det = re.search('(Intrusions detected: \d+),', line)
        int_prev = re.search('(Intrusions prevented: \d+)', line)
        if int_det:
            final_data.append(int_det.group(1))
        if int_prev:
            final_data.append(int_prev.group(1))
        if int_det and int_prev:
            break
    print(f"IPS report:\n{final_data}\n"+"-"*40)

def GAV(temp_file="GAV.txt"):
    final_data = []
    # Read data into variable
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        vir_det = re.search('(Virus detected: \d+)', line)
        if vir_det:
            final_data.append(vir_det.group(1))
        if vir_det:
            break
    print(f"GAV report:\n{final_data}\n"+"-"*40)

def botnet_trend(temp_file="botnet_trend.txt"):
    final_data = []
    # Read data into variable
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        src_ip = re.search('(Source IP blocked: \d+)', line)
        dest_ip = re.search('(Destination IP blocked: \d+)', line)
        if src_ip:
            final_data.append(src_ip.group(1))
        if dest_ip:
            final_data.append(dest_ip.group(1))
        if src_ip and dest_ip:
            break
    print(f"Botnet Activity Trend:\n{final_data}\n"+"-"*40)

def botnet_dest(temp_file="botnet_detect_dest.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)\n"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    i = 0
    for line in lines:
        if x == True:
            y = True
        if line == text:
            x = True
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
            i += 1
            if i > 2:
                break
    print("Botnet Detection by Destination")
    print(f"Destination: {final_data[0]} with {final_data[1]} hits @ {final_data[2]}%\n" + "-" * 40)

def block_botnet_sites(temp_file="botnet_block.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)\n"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if y == True:
            if line.__contains__("Total"):
                break
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line == text:
                x = True
            if x == True:
                y = True
    total_sites = len(final_data) // 3
    print("Blocked Botnet Sites")
    i = 0
    for site in range(total_sites):
        print(f"Source: {final_data[0 + i]} with {final_data[1 + i]} hits @ {final_data[2 + i]}%")
        i += 3
    print("-" * 40)

########################################################################################

def reports(temps):
    if "IPS.txt" in temps:
        IPS()
    if "GAV.txt" in temps:
        GAV()
    if "botnet_trend.txt" in temps:
        botnet_trend()
    if "botnet_detect_dest.txt" in temps:
        botnet_dest()
    if "botnet_block.txt" in temps:
        block_botnet_sites()
### MAIN ###
client_files = dir_list()
paths_in_dict(client_files)
temps = create_temps()
reports(temps)

### END ###