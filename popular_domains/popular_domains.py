from PyPDF2 import PdfFileReader
import re
import dns
import dns.resolver
import os

### VARIABLES ###
client_list = [['Intermax','0002.txt', '0002'],['Knudtsen','0005.txt', '0005'],['BankCDA','0007.txt', '0007'],['HONI','0021.txt', '0021'],['MMCO','0030.txt', '0030'],['Integrated Personnel','0038.txt', '0038'],['Northcon','0046.txt', '0046'],['Bayshore','0050.txt', '0050'],['MPMS','0073.txt', '0073'],['PFFM','0100.txt', '0100']]

### FUNCTIONS ###
def select_client(client_list):
    i = 1
    for item in client_list:
        print(f'{i}) {item[0]}')
        i += 1
    client_choice = int(input("Please enter a client selection by number:\n>")) - 1
    return client_list[client_choice][1], client_list[client_choice][2]

# Get list of files to choose from
def pdf_paths():                                                                # Create list of file paths for the pdfs
    files = os.listdir()                                                         # Get list of all files in the directory
    environment = os.getcwd() + "\\"                                            # Get local directory path
    pdfs = []
    for item in files:
        possible_pdf = environment + item                                            # Change the list entries to their full file path
        if possible_pdf[-3:] == "pdf":
            pdfs.append(possible_pdf)
    return pdfs

# Create list with IPs
def find_IPs(ip_list):
    ip_pattern = re.compile(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})')
    rfile = open("temp_nslookup.txt", "r")
    Lines = rfile.readlines()
    for line in Lines:
        each_line = line.strip()
        if ip_pattern.search(each_line) and not each_line.__contains__(client_num):
            ip_list.append(each_line)
    rfile.close()

# Create list with URLs
def create_URL_list(url_list, client_name):
    url_is = re.compile(r'[0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*(:(0-9)*)*(\/?)([a-zA-Z0-9\-\.\?\,\'\/\\\+&amp;%\$#_]*)?$')
    url_not = re.compile(r'[0-9]([-.\w]*[0-9])*((0-9)*)*(\/?)([0-9\.\?\,\'\/\\\+&amp;%\$#_]*)?$')

    rfile = open("temp_nslookup.txt", "r")
    Lines = rfile.readlines()
    for line in Lines:
        each_line = line.strip()
        if url_is.search(each_line) and not url_not.search(each_line) and not each_line.__contains__("cloudbackup") and not each_line.__contains__("zoom") and not each_line.__contains__("us-west") and not each_line.__contains__("gvt1") and not each_line.__contains__("cloudfront") and not each_line.__contains__("office") and not each_line.__contains__("logicnow") and not each_line.__contains__("n-able") and not each_line.__contains__("bitdefender") and not each_line.__contains__("google") and not each_line.endswith("...") and each_line != "Hits" and each_line != "MB" and not each_line.startswith("Domain") and each_line != "From:" and each_line != "To:" and not each_line.__contains__("microsoft") and not each_line.__contains__("breck") and not each_line.__contains__("windowsupdate") and not each_line.__contains__("quickbooks") and not each_line.__contains__("adobe") and not each_line.__contains__("verisign") and not each_line.__contains__("ocsp") and not each_line.startswith("Most") and not each_line.__contains__(client_name):
        #if url.search(each_line):
            url_list.append(each_line)
    rfile.close()


# Take a given URL and run nslookup
def get_ip(site, site_ips, site_dict):
    result = dns.resolver.resolve(site, 'A')
    for ipval in result:
        site_ips.append(ipval.to_text())
    site_dict[site] = site_ips

# run nslookup and add all IP address to the dictionary
def add_URLs_to_dict(url_list, site_dict):
    for site in url_list:
        site_ips = []
        try:
            get_ip(site, site_ips, site_dict)
        except:
            continue

# Add IP address to the dictionary
def add_IPs_to_dict(ip_list, site_dict):
    ip_list = set(ip_list)
    ip_list = list(ip_list)
    i = 1
    for address in ip_list:
        site_dict[f"IP {i}"] = address
        i += 1

def run_per_pdf(in_file, client_file, client_name):
    site_dict = {}
    ip_list = []
    url_list = []

    temp = open(in_file, 'rb')
    PDF_read = PdfFileReader(temp)
    num_pages = PDF_read.getNumPages()

    f = open("temp_nslookup.txt", "a")

    # Extract text and write to text file
    for page in range(num_pages):
        numbered_page = PDF_read.getPage(page)
        page_text = numbered_page.extractText()
        f.write(page_text)
    f.close()

    # Remove empty lines
    with open("temp_nslookup.txt", "r") as filehandle:
        lines = filehandle.readlines()
    with open("temp_nslookup.txt", 'w') as filehandle:
        lines = filter(lambda x: x.strip(), lines)
        filehandle.writelines(lines)
    filehandle.close()
    find_IPs(ip_list)
    create_URL_list(url_list, client_name)
    add_URLs_to_dict(url_list, site_dict)
    add_IPs_to_dict(ip_list, site_dict)

    os.remove("temp_nslookup.txt")

    f = open(client_file, "r")
    fwrite = open(client_file, "a")

    for key, value in site_dict.items():
        with open(client_file) as f:
            if key.startswith('IP') and value in f.read():
                continue
            elif key in f.read():
                continue
            else:
                print(f'{key}: {value}')
                if "IP" in key:
                    fwrite.write(f'{value}\n')
                else:
                    fwrite.write(f'{key}\n')
    f.close()
    fwrite.close()

def main_loop(file_paths, client_name):
    while True:
        i = 1
        for item in file_paths:
            print(f"{i}) {item}")
            i += 1
        chosen_file = input("Please select which file you would like to parse (press enter to exit):\n")
        if chosen_file == "":
            quit()
        # try:
        chosen_file = int(chosen_file)
        if chosen_file in range(1, i):
            run_per_pdf(file_paths[chosen_file - 1], client_file, client_name)
        elif chosen_file > i-1:
            print("That is not one of the options.")
        # except:
        #     print("Please enter only numbers.")

### MAIN ###
# Query user for the client to work with
client = select_client(client_list)
client_file = client[0]
client_num = client[1]

# Create list of pdf files to choose from
file_paths = pdf_paths()

# Run the main loop
main_loop(file_paths, client_num)


### END ###