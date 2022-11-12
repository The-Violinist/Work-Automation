# Commenting needed!

# -*- coding: utf-8 -*-

from datetime import date, timedelta
from os import listdir, remove
from PyPDF2 import PdfFileReader
import os
import re

### VARIABLES ###
# Dictionary of keywords to search for files in the weekly directories
f_type = {
    "app_use_bw": [
        "Application_Usage_Applications_(Host)_by_Bandwidth",
        "Application_Usage_Applications__Host__by_Bandwidth",
    ],
    "app_use_hits": [
        "Application_Usage_Applications_(Host)_by_Hits",
        "Application_Usage_Applications__Host__by_Hits",
        "Application_Usage_Top_Hosts_by_Hits",
    ],
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
    "top_cli_user": [
        "Top_Clients_Users__Sent_and_Received__by_Bandwidth",
        "Top_Clients_Users_(Sent_and_Received)_by_Bandwidth",
    ],
    "top_cli_host": [
        "Top_Clients_Hosts__Sent_and_Received__by_Bandwidth",
        "Top_Clients_Hosts_(Sent_and_Received)_by_Bandwidth",
    ],
    "top_cli_hits": ["Top_Clients_Hosts_by_Hits"],
    "IPS": ["Intrusions_(IPS)_Activity_Trend", "Intrusions__IPS__Activity_Trend"],
    "GAV": ["Virus_(GAV)_Activity_Trend", "Virus_Activity_Trend"],
    "proxy_bw": ["Proxy_Traffic_Destination_by_Bandwidth"],
    "proxy_hits": ["Proxy_Traffic_Destination_by_Hits"],
}

clients = [
    "Intermax",
    "Knudtsen",
    "BankCDA",
    "HONI",
    "MMCO",
    "Integrated Personnel",
    "Northcon",
    "Bay Shore",
    "PFFM",
]

### FUNCTIONS ###

#############################
# Gather files to work with #
#############################

# Select which client
def select_client():
    i = 1
    for item in clients:
        print(f"{i}) {item}")  # Print out a list of clients to select from
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
    # Parent for all client directories
    dir_path = (
        "\\\\FS01\\MSP-SecReview\\weekly"  # Upper level directory for all client files
    )

    # dir_path = 'C:\\Users\\darmstrong\\Desktop\\script_test'                    # Test directory

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
    cli_dir_path = folders[
        client
    ]  # Concatenate a path to the specific weekly folder for that client
    file_list = listdir(
        cli_dir_path
    )  # Create a list of all the files in that directory
    for item in file_list:
        path = cli_dir_path + "\\" + item
        all_files.append(path)  # Append all file paths to the files list
    return all_files, clients[client]


def paths_in_dict(cli_files):
    for k in f_type:  # Each key in the search word dictionary
        for value in f_type[k]:  # Each value for each key
            for item in cli_files:  # Each file in the directory
                if (
                    value in item
                ):  # If the search word is in the file name, change the value in the dictionary to the filepath
                    f_type[k] = item


######################################################################################

########################
# Extract Data to TXTs #
########################

# Extract data from PDFs
def extract_data(in_file, out_file):
    temp = open(in_file, "rb")
    PDF_read = PdfFileReader(temp)
    num_pages = PDF_read.getNumPages()

    fw = open(out_file, "a")

    # Extract text and write to text file
    for page in range(num_pages):
        numbered_page = PDF_read.getPage(page)
        page_text = numbered_page.extractText()
        fw.write(page_text)
    fw.close()

    # Remove empty lines
    with open(out_file, "r") as filehandle:
        lines = filehandle.readlines()
    with open(out_file, "w") as filehandle:
        lines = filter(lambda x: x.strip(), lines)
        filehandle.writelines(lines)
    filehandle.close()


# Create all txt files
def create_temps():
    temps = []
    for k in f_type:
        in_file = f_type[
            k
        ]  # Input file is the filepath which was added to the dictionary
        out_file = (
            f"{k}.txt"  # Output file is the dictionary key with txt file extension
        )
        try:
            extract_data(
                in_file, out_file
            )  # If a filepath exists as a value in the dictionary, extract all text to a temp txt file
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
        int_det = re.search("(Intrusions detected: \d+),", line)
        int_prev = re.search("(Intrusions prevented: \d+)", line)
        if int_det:
            final_data.append(int_det.group(1))
        if int_prev:
            final_data.append(int_prev.group(1))
        if int_det and int_prev:
            break
    f.write(f"\nIPS report:\n{final_data}\n" + "-" * 40)


def GAV(temp_file="GAV.txt"):
    final_data = []
    # Read data into variable
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        vir_det = re.search("(Virus detected: \d+)", line)
        if vir_det:
            final_data.append(vir_det.group(1))
        if vir_det:
            break
    f.write(f"\nGAV report:\n{final_data}\n" + "-" * 40)


def botnet_trend(temp_file="botnet_trend.txt"):
    final_data = []
    # Read data into variable
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        src_ip = re.search("(Source IP blocked: \d+)", line)
        dest_ip = re.search("(Destination IP blocked: \d+)", line)
        if src_ip:
            final_data.append(src_ip.group(1))
        if dest_ip:
            final_data.append(dest_ip.group(1))
        if src_ip and dest_ip:
            break
    f.write(f"\nBotnet Activity Trend:\n{final_data[0]}\n{final_data[1]}\n" + "-" * 40)


def botnet_dest(temp_file="botnet_detect_dest.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    total_sites = int(len(final_data) / 3)
    f.write("\nBotnet Detection by Destination\n")
    i = 0
    for site in range(total_sites):
        f.write(
            f"{final_data[0 + i]}: {final_data[1 + i]} hits at {final_data[2 + i]}%\n"
        )
        i += 3
    f.write("-" * 40)


def botnet_cli(temp_file="botnet_detect_cli.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    total_sites = int(len(final_data) / 3)
    f.write("\nBotnet Detection by Client\n")
    i = 0
    for site in range(total_sites):
        f.write(
            f"{final_data[0 + i]}: {final_data[1 + i]} hits at {final_data[2 + i]}%\n"
        )
        i += 3
    f.write("-" * 40)


def block_botnet_sites(temp_file="botnet_block.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    # x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                #     x = True
                # if x == True:
                y = True
    total_sites = int(len(final_data) / 3)
    f.write("\nBlocked Botnet Sites\n")
    i = 0
    # for site in range(total_sites):
    #     f.write(
    #         f"{final_data[0 + i]}: {final_data[1 + i]} hits at {final_data[2 + i]}%\n"
    #     )
    while i < len(final_data):
        if final_data[i + 1].isdigit():
            f.write(f"{final_data[i]}\n")
            i += 3
        else:
            while True:
                if final_data[i].__contains__(text):
                    i += 1
                    break
                else:
                    i += 1

    f.write("-" * 40)


def pop_domain_bytes(temp_file="pop_domain_bytes.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nPopular Domains by Bytes\n")
    i = 0
    for domain in range(3):
        if final_data[i + 3].isdigit():
            f.write(
                f"{final_data[i]}: {round((float(final_data[i + 1]) / 1024), 2)} GB at {final_data[i + 2]}%\n"
            )
            i += 5
        else:
            f.write(
                f"{final_data[i]}{final_data[i + 1]}: {round((float(final_data[i + 2]) / 1024), 2)} GB at {final_data[i + 3]}%\n"
            )
            i += 6
    f.write("-" * 40)


def pop_domain_hits(temp_file="pop_domain_hits.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nPopular Domains by Hits\n")
    i = 0
    for domain in range(3):
        if final_data[i + 3].isdigit():
            hits = "{:,}".format(
                int(final_data[i + 3])
            )  # Format the hits integer to use commas
            f.write(
                f"{final_data[i]}: {hits} hits at {final_data[i + 4]}%\n"
            )  # Print in the format: Domain, hits, percent
            i += 5  # Increment the index for the next domain
        else:
            hits = "{:,}".format(
                int(final_data[i + 4])
            )  # Format the hits integer to use commas
            f.write(
                f"{final_data[i]}{final_data[i + 1]}: {hits} hits at {final_data[i + 5]}%\n"
            )  # Print in the format: Domain, hits, percent
            i += 6  # Increment the index for the next domain
    f.write("-" * 40)


def proxy_bw(temp_file="proxy_bw.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Page 1"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nProxy Traffic by Bandwidth\n")
    i = 0
    for domain in range(1):
        if final_data[i + 3].isdigit():
            f.write(
                f"Destination by Bandwidth: {final_data[i]} with {round((float(final_data[i + 1]) / 1024), 2)} GB at {final_data[i + 2]}%\n"
            )
            i += 5
        else:
            f.write(
                f"Destination by Bandwidth: {final_data[i]}{final_data[i + 1]} with {round((float(final_data[i + 2]) / 1024), 2)} GB at {final_data[i + 3]}%\n"
            )
            i += 6
    f.write("-" * 40)


def proxy_hits(temp_file="proxy_hits.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Page 1"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nProxy Traffic by Hits\n")
    i = 0
    for domain in range(1):
        if final_data[i + 3].isdigit():
            hits = "{:,}".format(
                int(final_data[i + 3])
            )  # Format the hits integer to use commas
            f.write(
                f"Destination by Hits: {final_data[i]} with {hits} hits at {final_data[i + 4]}%\n"
            )  # Print in the format: Domain, hits, percent
            i += 5  # Increment the index for the next domain
        else:
            hits = "{:,}".format(
                int(final_data[i + 4])
            )  # Format the hits integer to use commas
            f.write(
                f"Destination by Hits: {final_data[i]}{final_data[i + 1]} with {hits} hits at {final_data[i + 5]}%\n"
            )  # Print in the format: Domain, hits, percent
            i += 6  # Increment the index for the next domain
    f.write("-" * 40)


def top_cli_host(temp_file="top_cli_host.txt"):
    # Get the text for the last string before the needed data
    text = "(%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    indexIncrementer = 5
    for line in lines:
        if line.__contains__("Hostname"):
            indexIncrementer = 6
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nTop Client hosts by Bytes\n")
    i = 0
    if indexIncrementer == 5:
        for host in range(3):
            f.write(
                f"{final_data[i]}: {round((float(final_data[i + 3]) / 1024), 2)} GB at {final_data[i + 4]}%\n"
            )
            i += indexIncrementer
    else:
        for host in range(3):
            if float(final_data[i + 5]) < 100:
                f.write(
                    f"{final_data[i + 1]} ({final_data[i]}): {round((float(final_data[i + 4]) / 1024), 2)} GB at {final_data[i + 5]}%\n"
                )
                i += indexIncrementer
            elif final_data[i + 1][0].isdigit():
                f.write(
                    f"{final_data[i + 1]}{final_data[i + 2]} ({final_data[i]}): {round((float(final_data[i + 5]) / 1024), 2)} GB at {final_data[i + 6]}%\n"
                )
                i += indexIncrementer + 1
            else:
                f.write(
                    f"{final_data[i + 2]} ({final_data[i]}{final_data[i + 1]}): {round((float(final_data[i + 5]) / 1024), 2)} GB at {final_data[i + 6]}%\n"
                )
                i += indexIncrementer + 1
    f.write("-" * 40)


def top_cli_user(temp_file="top_cli_user.txt"):
    # Get the text for the last string before the needed data
    text = "Bytes (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nTop Clients Users by Bandwidth\n")
    total_users = int(len(final_data) / 5)
    if total_users > 3:
        total_users = 3
    i = 0
    for user in range(total_users):
        if final_data[i + 1][0].isalpha():
            f.write(
                f"{final_data[i]}{final_data[i + 1]}: {round((float(final_data[i + 4]) / 1024), 2)} GB at {final_data[i + 5]}%\n"
            )
            i += 6
        else:
            f.write(
                f"{final_data[i]}: {round((float(final_data[i + 3]) / 1024), 2)} GB at {final_data[i + 4]}%\n"
            )
            i += 5
    f.write("-" * 40)


def top_cli_hits(temp_file="top_cli_hits.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    indexIncrement = 3
    for line in lines:
        if line.__contains__("Hostname"):
            indexIncrement = 4
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nTop Clients Hosts by Hits\n")
    i = 0
    if indexIncrement == 4:
        for host in range(3):
            if final_data[i + 2].isdigit():
                if int(final_data[i + 2]) < 1000000:
                    hits = "{:,}".format(
                        int(final_data[i + 2])
                    )  # Format the hits integer to use commas
                else:
                    hits = f"{round((float(final_data[i + 2]) / 1000000), 2)} million"
                f.write(
                    f"{final_data[i + 1]} ({final_data[i]}): {hits} hits at {final_data[i + 3]}%\n"
                )  # Print in the format: IP (hostname), hits, percent
                i += 4  # Increment the index for the next domain
            else:
                if int(final_data[i + 3]) < 1000000:
                    hits = "{:,}".format(
                        int(final_data[i + 3])
                    )  # Format the hits integer to use commas
                else:
                    hits = f"{round((float(final_data[i + 3]) / 1000000), 2)} million"
                f.write(
                    f"{final_data[i + 2]} ({final_data[i]}{final_data[i + 1]}): {hits} hits at {final_data[i + 4]}%\n"
                )  # Print in the format: IP (hostname), hits, percent
                i += 5  # Increment the index for the next domain
    else:
        for host in range(3):
            hits = "{:,}".format(
                int(final_data[i + 1])
            )  # Format the hits integer to use commas
            f.write(
                f"{final_data[i]}: {hits} hits at {final_data[i + 2]}%\n"
            )  # Print in the format: IP (hostname), hits, percent
            i += 3  # Increment the index for the next domain

    f.write("-" * 40)


def active_cli_bw(temp_file="active_client_bw.txt"):
    # Get the text for the last string before the needed data
    x = False
    y = False
    hits = 0
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    indexIncrement = 4
    for line in lines:
        if line.__contains__("Host"):
            indexIncrement = 5
        if line.__contains__("Hits\n"):
            hits += 1
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if hits == 3:
                x = True
            if x == True:
                y = True
    f.write("\nMost Active Clients by Bandwidth\n")
    i = 0
    if indexIncrement == 5:
        for host in range(3):
            if final_data[i + 4].isdigit():
                f.write(
                    f"{final_data[i + 2]} ({final_data[i + 1]}): {round((float(final_data[i + 3]) / 1024), 2)} GB\n"
                )
                i += 5
            else:
                f.write(
                    f"{final_data[i + 3]} ({final_data[i + 1]}{final_data[i + 2]}): {round((float(final_data[i + 4]) / 1024), 2)} GB\n"
                )
                i += 6
    else:
        for host in range(3):
            f.write(
                f"{final_data[i + 1]}: {round((float(final_data[i + 2]) / 1024), 2)} GB\n"
            )
            i += 4
    f.write("-" * 40)


def active_cli_hits(temp_file="active_client_hit.txt"):
    # Get the text for the last string before the needed data
    x = False
    y = False
    hits_text = 0
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    indexIncrement = 4
    for line in lines:
        if line.__contains__("Host"):
            indexIncrement = 5
        for line in lines:
            if line.__contains__("Hits"):
                hits_text += 1
            if line.__contains__("Total:"):
                break
            if y == True:
                each_line = line.strip()
                final_data.append(each_line)
            else:
                if hits_text == 4:
                    x = True
                if x == True:
                    y = True
    f.write("\nMost Active Clients by Hits\n")
    i = 0
    if indexIncrement == 5:
        for host in range(3):
            if final_data[i + 4].isdigit():
                hits_num = "{:,}".format(
                    int(final_data[i + 4])
                )  # Format the hits integer to use commas
                f.write(
                    f"{final_data[i + 2]} ({final_data[i + 1]}): {hits_num} hits\n"
                )  # Print in the format: Domain, hits, percent
                i += 5  # Increment the index for the next domain
            else:
                hits_num = "{:,}".format(
                    int(final_data[i + 5])
                )  # Format the hits integer to use commas
                f.write(
                    f"{final_data[i + 3]} ({final_data[i + 1]}{final_data[i + 2]}): {hits_num} hits\n"
                )  # Print in the format: Domain, hits, percent
                i += 6  # Increment the index for the next domain
    else:
        for host in range(3):
            hits_num = "{:,}".format(
                int(final_data[i + 3])
            )  # Format the hits integer to use commas
            f.write(
                f"{final_data[i + 1]} – {hits_num} hits\n"
            )  # Print in the format: Domain, hits, percent
            i += 4  # Increment the index for the next domain
    f.write("-" * 40)


def app_use_bw(temp_file="app_use_bw.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nApplication Usage by Bandwidth\n")
    i = 0
    for app in range(3):
        try:
            f.write(
                f"{final_data[i]}: {round((float(final_data[i + 1]) / 1024), 2)} GB at {final_data[i + 2]}%\n"
            )
            i += 5
        except:
            f.write(
                f"{final_data[i]} {final_data[i + 1]}: {round((float(final_data[i + 2]) / 1024), 2)} GB at {final_data[i + 3]}%\n"
            )
            i += 6
    f.write("-" * 40)


def app_use_hits(temp_file="app_use_hits.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nApplication Usage by Hits\n")
    i = 0
    for app in range(3):
        if final_data[i + 3].isdigit():
            if int(final_data[i + 3]) < 1000000:
                hits = "{:,}".format(
                    int(final_data[i + 3])
                )  # Format the hits integer to use commas
            else:
                hits = f"{round((float(final_data[i + 3]) / 1000000), 2)} million"
            f.write(
                f"{final_data[i]}: {hits} hits at {final_data[i + 4]}%\n"
            )  # Print in the format: app, hits, percent
            i += 5  # Increment the index for the next app
        else:
            if int(final_data[i + 4]) < 1000000:
                hits = "{:,}".format(
                    int(final_data[i + 4])
                )  # Format the hits integer to use commas
            else:
                hits = f"{round((float(final_data[i + 4]) / 1000000), 2)} million"
            f.write(
                f"{final_data[i]}{final_data[i + 1]}: {hits} hits at {final_data[i + 5]}%\n"
            )  # Print in the format: app, hits, percent
            i += 6  # Increment the index for the next app
    f.write("-" * 40)


def block_sites_cat(temp_file="block_sites_cat.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nBlocked Sites by Category\n")
    i = 0  # Incrementer for the rows of data
    num_entries = int(len(final_data) / 3)
    if num_entries > 3:
        num_entries = 3
    for entry in range(num_entries):
        if final_data[i + 1].isdigit():
            hits = "{:,}".format(
                int(final_data[1 + i])
            )  # Format the hits integer to use commas
            f.write(
                f"{final_data[0 + i]}: {hits} hits at {final_data[2 + i]}%\n"
            )  # Print in the format: Domain, hits, percent
            i += 3  # Increment the index for the next domain
        else:
            hits = "{:,}".format(
                int(final_data[i + 2])
            )  # Format the hits integer to use commas
            f.write(
                f"{final_data[i]}{final_data[i + 1]}: {hits} hits at {final_data[i + 3]}%\n"
            )  # Print in the format: Domain, hits, percent
            i += 4  # Increment the index for the next domain
    f.write("-" * 40)


def block_sites_cli(temp_file="block_sites_cli.txt"):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    os.remove(temp_file)
    for line in lines:
        if line.__contains__("Total:"):
            break
        if y == True:
            each_line = line.strip()
            final_data.append(each_line)
        else:
            if line.__contains__(text):
                x = True
            if x == True:
                y = True
    f.write("\nBlocked Sites by Client\n")
    i = 0  # Incrementer for the rows of data
    num_entries = int(len(final_data) / 3)
    if num_entries > 3:
        num_entries = 3
    for entry in range(num_entries):
        hits = "{:,}".format(
            int(final_data[1 + i])
        )  # Format the hits integer to use commas
        f.write(
            f"{final_data[0 + i]}: {hits} hits at {final_data[2 + i]}%\n"
        )  # Print in the format: Domain, hits, percent
        i += 3  # Increment the index for the next domain
    f.write("-" * 40)


########################################################################################


def reports(temps):
    if "top_cli_host.txt" in temps:
        top_cli_host()
    if "top_cli_user.txt" in temps:
        top_cli_user()
    if "top_cli_hits.txt" in temps:
        top_cli_hits()
    if "app_use_bw.txt" in temps:
        app_use_bw()
    if "app_use_hits.txt" in temps:
        app_use_hits()
    if "block_sites_cat.txt" in temps:
        block_sites_cat()
    if "block_sites_cli.txt" in temps:
        block_sites_cli()
    if "pop_domain_bytes.txt" in temps:
        pop_domain_bytes()
    if "pop_domain_hits.txt" in temps:
        pop_domain_hits()
    if "proxy_bw.txt" in temps:
        proxy_bw()
    if "proxy_hits.txt" in temps:
        proxy_hits()
    if "botnet_trend.txt" in temps:
        botnet_trend()
    if "botnet_detect_dest.txt" in temps:
        botnet_dest()
    if "botnet_detect_cli.txt" in temps:
        botnet_cli()
    if "botnet_block.txt" in temps:
        block_botnet_sites()
    if "IPS.txt" in temps:
        IPS()
    if "GAV.txt" in temps:
        GAV()
    if "active_client_bw.txt" in temps:
        active_cli_bw()
    if "active_client_hit.txt" in temps:
        active_cli_hits()


### MAIN ###
client_files = dir_list()
f = open(f"{client_files[1]}_WG_data.txt", "a")
paths_in_dict(client_files[0])
temps = create_temps()
reports(temps)
f.close()
### END ###
