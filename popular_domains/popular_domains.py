from PyPDF2 import PdfFileReader
import re
import dns
import dns.resolver
import os

in_file = "test3.pdf"
site_dict = {}
ip_list = []
url_list = []

temp = open(in_file, 'rb')
PDF_read = PdfFileReader(temp)
num_pages = PDF_read.getNumPages()

f = open("test.txt", "a")

# Extract text and write to text file
for page in range(num_pages):
    numbered_page = PDF_read.getPage(page)
    page_text = numbered_page.extractText()
    f.write(page_text)
f.close()

# Remove empty lines
with open("test.txt", "r") as filehandle:
    lines = filehandle.readlines()
with open("test.txt", 'w') as filehandle:
    lines = filter(lambda x: x.strip(), lines)
    filehandle.writelines(lines)
filehandle.close()

# Create list with IPs
def find_IPs():
    ip_pattern = re.compile(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})')
    rfile = open("test.txt", "r")
    Lines = rfile.readlines()
    for line in Lines:
        each_line = line.strip()
        if ip_pattern.search(each_line) and not each_line.__contains__("0038"):
            ip_list.append(each_line)
    rfile.close()

# Create list with URLs
def create_URL_list():
    url_is = re.compile(r'[0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*(:(0-9)*)*(\/?)([a-zA-Z0-9\-\.\?\,\'\/\\\+&amp;%\$#_]*)?$')
    url_not = re.compile(r'[0-9]([-.\w]*[0-9])*((0-9)*)*(\/?)([0-9\.\?\,\'\/\\\+&amp;%\$#_]*)?$')

    rfile = open("test.txt", "r")
    Lines = rfile.readlines()
    for line in Lines:
        each_line = line.strip()
        if url_is.search(each_line) and not url_not.search(each_line) and not each_line.__contains__("cloudbackup") and not each_line.__contains__("zoom") and not each_line.__contains__("us-west") and not each_line.__contains__("gvt1") and not each_line.__contains__("cloudfront") and not each_line.__contains__("office") and not each_line.__contains__("logicnow") and not each_line.__contains__("n-able") and not each_line.__contains__("bitdefender") and not each_line.__contains__("google") and not each_line.endswith("...") and each_line != "Hits" and each_line != "MB" and not each_line.startswith("Domain") and each_line != "From:" and each_line != "To:" and not each_line.__contains__("microsoft") and not each_line.__contains__("breck") and not each_line.__contains__("windowsupdate") and not each_line.__contains__("quickbooks") and not each_line.__contains__("adobe") and not each_line.__contains__("verisign") and not each_line.__contains__("ocsp") and not each_line.startswith("Most"):
        #if url.search(each_line):
            url_list.append(each_line)
    rfile.close()


# Take a given URL and run nslookup
def get_ip(site, site_ips):
    result = dns.resolver.resolve(site, 'A')
    for ipval in result:
        site_ips.append(ipval.to_text())
    site_dict[site] = site_ips

# run nslookup and add all IP address to the dictionary
def add_URLs_to_dict():
    for site in url_list:
        site_ips = []
        try:
            get_ip(site, site_ips)
        except:
            continue

# Add IP address to the dictionary
def add_IPs_to_dict(ip_list):
    ip_list = set(ip_list)
    ip_list = list(ip_list)
    i = 1
    for address in ip_list:
        site_dict[f"IP {i}"] = address
        i += 1


### MAIN ###
find_IPs()
create_URL_list()
add_URLs_to_dict()
add_IPs_to_dict(ip_list)

os.remove("test.txt")

for key, value in site_dict.items():
    print(key, value)
# num = 1
# # print(ip_list)
# for item in url_list:
#     print(f"{num}. {item}")
#     num += 1