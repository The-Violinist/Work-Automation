import dns
import dns.resolver
from datetime import date, timedelta
from os import listdir, remove, system
from PyPDF2 import PdfFileReader
import re

### VARIABLES ###
# Dictionary of keywords to search for files in the weekly directories
f_type = {
    "pop_domain_bytes": ["Most_Popular_Domains_Bytes"],
    "pop_domain_hits": ["Most_Popular_Domains_Hits"],
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
        # "C:\\Users\\darmstrong\\Desktop\\script_test"
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


def pop_domains(ip_list, temp_file):
    # Get the text for the last string before the needed data
    text = "Hits (%)"
    x = False
    y = False
    final_data = []
    with open(temp_file, "r") as read_file:
        lines = read_file.readlines()
    read_file.close()
    remove(temp_file)
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
    i = 0
    for domain in range(50):
        fRead = open("test.txt", "r")
        if final_data[i + 3].isdigit():
            if f"{final_data[i]}" not in fRead.read():
                ip_list.append(f"{final_data[i]}")
            i += 5
        elif final_data[i + 4].isdigit():
            if f"{final_data[i]}{final_data[i + 1]}" not in fRead.read():
                ip_list.append(f"{final_data[i]}{final_data[i + 1]}")
            i += 6
        else:
            if (
                f"{final_data[i]}{final_data[i + 1]}{final_data[i + 2]}"
                not in fRead.read()
            ):
                ip_list.append(f"{final_data[i]}{final_data[i + 1]}{final_data[i + 2]}")
            i += 7
        fRead.close()
        if domain == 49:
            break
        if (
            final_data[i + 3].isdigit()
            or final_data[i + 4].isdigit()
            or final_data[i + 5].isdigit()
        ):
            continue
        else:
            while True:
                if final_data[i].__contains__(text):
                    i += 1
                    break
                else:
                    i += 1


def url_list():
    ip_list = []
    client_files = dir_list()
    paths_in_dict(client_files[0])
    temps = create_temps()
    if "pop_domain_bytes.txt" in temps:
        pop_domains(ip_list, "pop_domain_bytes.txt")
    if "pop_domain_hits.txt" in temps:
        pop_domains(ip_list, "pop_domain_hits.txt")
    return ip_list


########################################################################################

################
# Analyze Data #
################

# Take a given URL and run nslookup
def get_ip(site, site_ips, site_dict):
    result = dns.resolver.resolve(site, "A")
    for ipval in result:
        site_ips.append(ipval.to_text())
    site_dict[site] = site_ips


def is_ip(site):
    match_obj = re.search(r"^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})", site)
    if match_obj is None:
        return False
    else:
        for value in match_obj.groups():
            if int(value) > 255:
                return False
        site_dict[site] = [site]
    return site


# run get_ip and add all IP address to the dictionary
def add_urls_to_dict(url_list, site_dict):
    for site in url_list:
        if is_ip(site) == False:
            try:
                site_ips = []
                get_ip(site, site_ips, site_dict)
            except:
                continue


### MAIN ###
site_dict = {}
urls = url_list()
add_urls_to_dict(urls, site_dict)

fWrite = open("test.txt", "a")

for k, v in site_dict.items():
    print(k, v)
    while True:
        response = input("Please enter a selection:\n1) Safe\n2)Malicious\n>")
        if response == "1":
            fWrite.write(f"{k}\n")
            break
        elif response == "2":
            break
        else:
            print('Please enter either "1" or "2"')
    system("cls")
fWrite.close()
### END ###
