import xlsxwriter
import requests
import json
import urllib3
from getpass import getpass
from os import system

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# List of sites in the Unifi controller
all_sites = {
    "37": "0037 Triple Play",
    "31": "0031 North Idaho Dermatology",
    "2": "0002 - Intermax Networks",
    "2a": "0002 - Intermax Lone Mtn",
    "2b": "0002 - Intermax RVT Conf",
    "5a": "0005 Knudtsen Abra",
    "18": "0018 CDA Honda",
    "30": "0030 Magnuson McHugh",
    "45": "0045 JVW Law",
    "14": "0014 North Idaho CASA",
    "40": "0040 Smith and Malek",
    "2c": "0002 - Innovation Den",
    "38": "0038 Integrated Personnel",
    "46": "0046 Northcon",
    "50": "0050 Bayshore Systems",
    "11": "0011 Tobler",
    "51": "0051 Innercept",
    "47": "0047 Shabby",
    "59": "0059 Habitat",
    "72": "0072 HeaterCraft",
    "5": "0005 Knudtsen Main",
    "54": "0054 Total Tax",
    "61": "0061 Kuma",
    "63": "0063 Dyck's Oil",
    "64": "0064 Widmyer Corporation",
    "20": "0020 T&S",
    "42": "0042 HMH",
    "34": "0034 ITS",
    "21": "0021 Hospice of North Idaho",
    "67": "0067 NNCE",
    "68": "0068 - Aaging Better",
    "52": "0052 - ABHS",
    "75": "TP4346 - Caddyshack",
    "15": "0015 - Northern Management Services",
    "12": "0012 - Stancraft",
    "74": "0074 - St. Pius",
    "91": "0091 - Featherston",
    "12a": "0012 Stancraft Import",
    "76": "0076 - EmpireEye",
    "16": "0016 - Community Title",
    "1": "0001 - ADB",
    "80": "0080 - BookWorks",
    "81": "0081 Tesh",
    "82": "0082 - Paul Daugharty",
    "85": "0085-AuburnCrest",
    "7": "0007 - Bankcda",
    "17": "0017 HADDOCK",
    "88": "0088-BGC",
    "19": "0019 - CDAEDC",
    "86": "0086 - Owsley Plastic Surgery",
    "62": "0062 - Association Services",
    "35": "0035 Timberline",
    "84": "0084 - Inland NW Spine",
    "95": "0095 Boundary Road & Bridge",
    "96": "0096 Lake Drive Apartments",
    "97": "0097-MarkJackson",
    "93": "0093-CDA PEDS",
    "27": "0027 - Callahan",
    "100": "0100 PFFM",
    "106": "0106 Silver Pine Wealth Management",
    "2d": "0002-WhalenResidence",
    "108": "0108 - Hayden Canyon Charter School",
    "69": "0069 - Atchley Financial",
    "41": "0041 CDASSE",
    "113": "0113 - Ability Home Health",
    "114": "0114 GarageSkins.com",
    "115": "0115-BrandIt",
    "119": "0119 Summit Rehab",
    "120": "0120 City of Post Falls",
    "105": "0105 - Holy Family Catholic School",
    "10": "0010 - Minutepress",
    "124": "0124 -Family Health Center",
    "125": "0125 - Killer Burger",
    "43": "0043 - SwissTech",
    "130": "0130-GVD-SteamPlant",
    "102": "0102-Rainbow Electric",
    "133": "0133 - The Timbers",
    "132": "0132 Specialty Construction",
    "134": "0134 - 1250 Ironwood",
    "145": "0145-ASC Northwest",
    "138": "0138-Calvary",
    "135": "0135-RubyOnTheRiver",
    "137": "0137-Ruby Suites",
    "147": "0147-Sravasti Abbey",
    "151": "0151-Childrens Village",
    "36": "0036-Associated Credit",
    "149": "0149-Steam Plant Hotel",
    "152": "0152-Hotel Ruby-Ponderay",
    "148": "0148-Hotel Ruby-Spokane",
    "150": "0150-Montvale Hotel",
}


#########
# EXCEL #
#########

### Create file headers and formatting ###

# Formatting for client devices
def client_start_file(client_col_lens):
    ### Formatting ###
    headers = [
        "Name",
        "MAC",
        "IP",
        "OUI",
        "Connection",
        "Access Point",
    ]
    header = wb.add_format(
        {"bold": True, "align": "center", "bg_color": "E0F4FF", "border": 1}
    )
    ws.freeze_panes(1, 0)

    i = 0
    for item in headers:
        ws.write(0, i, item, header)
        i += 1

    column_num = 0
    for item in client_col_lens:
        ws.set_column(column_num, column_num, client_col_lens[item])
        column_num += 1


def client_write_data(vals, row):
    general_cells = wb.add_format({"align": "left"})
    concerning_cells = wb.add_format({"align": "left", "bg_color": "F8FA9B"})
    column = 0
    for k, v in vals.items():
        if (column == 6 and v == True) or (column == 7 and v == False):
            ws.write(row, column, v, concerning_cells)
        else:
            ws.write(row, column, v, general_cells)
        column += 1


# Formatting for Unifi devices
def unifi_start_file(col_lens):
    ### Formatting ###
    headers = [
        "Name",
        "MAC",
        "IP",
        "Model",
        "Version",
        "Uplink",
        "EoL Status",
        "Connected",
    ]
    header = wb.add_format(
        {"bold": True, "align": "center", "bg_color": "E0F4FF", "border": 1}
    )
    ws.freeze_panes(1, 0)

    i = 0
    for item in headers:
        ws.write(0, i, item, header)
        i += 1

    column_num = 0
    for item in col_lens:
        ws.set_column(column_num, column_num, col_lens[item])
        column_num += 1


def unifi_write_data(vals, row):
    general_cells = wb.add_format({"align": "left"})
    concerning_cells = wb.add_format({"align": "left", "bg_color": "F8FA9B"})
    column = 0
    for k, v in vals.items():
        if (column == 6 and v == True) or (column == 7 and v == False):
            ws.write(row, column, v, concerning_cells)
        else:
            ws.write(row, column, v, general_cells)
        column += 1


# Minimum column size for Unifi device spreadsheet
unifi_col_lens = {
    "name": 6,
    "mac": 5,
    "ip": 4,
    "model": 7,
    "version": 9,
    "uplink": 8,
    "eol": 12,
    "online": 11,
}

# Minimum column size for client device spreadsheet
client_col_lens = {
    "name": 6,
    "mac": 5,
    "ip": 4,
    "oui": 5,
    "essid": 12,
    "uplink": 14,
}


#########################
# Unifi Controller Data #
#########################

# Raw data
def site_data(site_requested, username, password):
    # Arguments for session
    gateway = {"ip": "unifi.intermaxnetworks.com", "port": 8443}
    headers = {"Accept": "application/json", "Content-Type": "application.json"}
    loginUrl = "api/login"
    url = f"https://{gateway['ip']}:{gateway['port']}/{loginUrl}"
    body = {"username": username, "password": password}

    # Create session and login
    session = requests.Session()
    response = session.post(url, headers=headers, data=json.dumps(body), verify=False)
    api_data = response.json()

    # List of sites in controller
    getSitesUrl = "api/self/sites"
    url = f"https://{gateway['ip']}:{gateway['port']}/{getSitesUrl}"

    # Take list of sites and get the upper level data
    response = session.get(url, headers=headers, verify=False)
    api_data = response.json()
    responseList = api_data["data"]

    # Selects site based off of the Unifi supplied name
    for items in responseList:
        # If desc == the friendly name, get the Unifi name
        if items.get("desc") == all_sites[site_requested]:
            n = items.get("name")

    # Gather all data for the client devices
    getDevicesUrl = f"api/s/{n}/stat/sta"
    url = f"https://{gateway['ip']}:{gateway['port']}/{getDevicesUrl}"
    response = session.get(url, headers=headers, verify=False)
    api_data = response.json()
    responseList = api_data["data"]

    # Gather all data for the Unifi devices
    getUnifiDevicesUrl = f"api/s/{n}/stat/device"
    unifiUrl = f"https://{gateway['ip']}:{gateway['port']}/{getUnifiDevicesUrl}"
    unifiResponse = session.get(unifiUrl, headers=headers, verify=False)
    unifiApi_data = unifiResponse.json()
    unifiResponseList = unifiApi_data["data"]

    return responseList, unifiResponseList


# Gather the relevant data for Unifi devices
def unifi_device_data(raw_data):
    devices = {}
    total_devices = 1
    for device in raw_data:
        name = "Unknown"
        try:
            name = device["name"]
        except:
            pass
        mac = "Unknown"
        try:
            mac = device["mac"]
        except:
            pass
        model = "Unknown"
        try:
            model = device["model"]
        except:
            pass
        ip = "Unknown"
        try:
            ip = device["ip"]
        except:
            pass
        uplink = "Unknown"
        try:
            uplink = f"{device['uplink']['uplink_device_name']} #{device['uplink']['uplink_remote_port']}"
        except:
            pass
        eol = "Unknown"
        try:
            eol = device["model_in_eol"]
        except:
            pass
        version = "Unknown"
        try:
            version = device["version"]
        except:
            pass
        connected = False
        try:
            internet = device["start_connected_millis"]
            connected = True
        except:
            pass

        devices[str(total_devices)] = {
            "name": name,
            "mac": mac,
            "ip": ip,
            "model": model,
            "version": version,
            "uplink": uplink,
            "eol": eol,
            "online": connected,
        }
        total_devices += 1
    return devices


# Gather the relevant data for client devices
def client_device_data(raw_data, unifi_devices):
    devices = {}
    total_devices = 1
    for device in raw_data:
        name = "Unknown"
        try:
            name = device["hostname"]
        except:
            pass
        mac = "Unknown"
        try:
            mac = device["mac"]
        except:
            pass
        ip = "Unknown"
        try:
            ip = device["ip"]
        except:
            pass
        oui = "Unknown"
        try:
            oui = f"{device['oui']}"
            if oui == "":
                oui = "Unknown"
        except:
            pass
        essid = "Wired"
        try:
            essid = device["essid"]
        except:
            pass
        uplink = "Unknown"
        try:
            uplink = device["ap_mac"]
        except:
            pass
        if uplink == "Unknown":
            try:
                uplink = f"{device['sw_mac']} #{device['sw_port']}"
            except:
                pass
        if len(uplink) != 17:
            for device in unifi_devices:
                if unifi_devices[device]["mac"] == uplink[:17]:
                    uplink = f"{unifi_devices[device]['name']}{uplink[17:]}"
        elif len(uplink) == 17:
            for device in unifi_devices:
                if unifi_devices[device]["mac"] == uplink:
                    uplink = unifi_devices[device]["name"]
        devices[str(total_devices)] = {
            "name": name,
            "mac": mac,
            "ip": ip,
            "oui": oui,
            "essid": essid,
            "uplink": uplink,
        }
        total_devices += 1
    return devices


# Max cell width for client device cells
def client_cell_width(final_data):
    for device in final_data.items():
        if len(str(device[1]["name"])) + 2 > client_col_lens["name"]:
            client_col_lens["name"] = len(str(device[1]["name"])) + 2
        if len(str(device[1]["mac"])) + 2 > client_col_lens["mac"]:
            client_col_lens["mac"] = len(str(device[1]["mac"])) + 2
        if len(str(device[1]["ip"])) + 2 > client_col_lens["ip"]:
            client_col_lens["ip"] = len(str(device[1]["ip"])) + 2
        if len(str(device[1]["oui"])) + 2 > client_col_lens["oui"]:
            client_col_lens["oui"] = len(str(device[1]["oui"])) + 2
        if len(str(device[1]["essid"])) + 2 > client_col_lens["essid"]:
            client_col_lens["essid"] = len(str(device[1]["essid"])) + 2
        if len(str(device[1]["uplink"])) + 2 > client_col_lens["uplink"]:
            client_col_lens["uplink"] = len(str(device[1]["uplink"])) + 2


# Max cell width for Unifi device cells
def unifi_cell_width(final_data):
    for device in final_data.items():
        if len(str(device[1]["name"])) + 2 > unifi_col_lens["name"]:
            unifi_col_lens["name"] = len(str(device[1]["name"])) + 2
        if len(str(device[1]["mac"])) + 2 > unifi_col_lens["mac"]:
            unifi_col_lens["mac"] = len(str(device[1]["mac"])) + 2
        if len(str(device[1]["ip"])) + 2 > unifi_col_lens["ip"]:
            unifi_col_lens["ip"] = len(str(device[1]["ip"])) + 2
        if len(str(device[1]["model"])) + 2 > unifi_col_lens["model"]:
            unifi_col_lens["model"] = len(str(device[1]["model"])) + 2
        if len(str(device[1]["version"])) + 2 > unifi_col_lens["version"]:
            unifi_col_lens["version"] = len(str(device[1]["version"])) + 2
        if len(str(device[1]["uplink"])) + 2 > unifi_col_lens["uplink"]:
            unifi_col_lens["uplink"] = len(str(device[1]["uplink"])) + 2
        if len(str(device[1]["eol"])) + 2 > unifi_col_lens["eol"]:
            unifi_col_lens["eol"] = len(str(device[1]["eol"])) + 2
        if len(str(device[1]["online"])) + 2 > unifi_col_lens["online"]:
            unifi_col_lens["online"] = len(str(device[1]["online"])) + 2


# Take all data and create client devices spreadsheet
def client_devices():
    try:
        # Unifi Device data -- Names and MACs
        unifi_raw_data = site_data(site_requested, username, password)[1]
        unifi_final_data = unifi_device_data(unifi_raw_data)

        # Client device data
        raw_data = site_data(site_requested, username, password)[0]
        final_data = client_device_data(raw_data, unifi_final_data)
        client_cell_width(final_data)
        client_start_file(client_col_lens)

        row = 1
        for k, v in final_data.items():
            client_write_data(v, row)
            row += 1

        wb.close()
    except:
        print("Unable to create file.")
        input("Press any key to continue...")
        system("cls")


# Take all data and create Unifi devices spreadsheet
def unifi_devices():
    try:
        unifi_raw_data = site_data(site_requested, username, password)[1]
        final_data = unifi_device_data(unifi_raw_data)
        unifi_cell_width(final_data)
        unifi_start_file(unifi_col_lens)

        row = 1
        for k, v in final_data.items():
            unifi_write_data(v, row)
            row += 1

        wb.close()
    except:
        print("Unable to create file.")
        input("Press any key to continue...")
        system("cls")


########
# MAIN #
########
print("NOTE: input will not appear on screen.")
username = getpass("Enter Unifi username:\n>")
password = getpass("Enter Unifi password:\n>")

# Loop until the user decides to quit
while True:
    # Request a site
    site_requested = str(
        input(
            "Enter a client number without leading zeros (e.g Intermax = 2) or press enter to exit:\n>"
        )
    )
    # User can exit the program by pressing enter with no input
    if site_requested == "":
        quit()

    # Verify input
    if site_requested not in all_sites:
        print(f"{site_requested} is not a valid client number")
        input("Press any key to continue...")
        system("cls")
        continue

    # User selects either Unifi or client device data
    while True:
        list_type = input(
            "Make a selection:\n1) Unifi Device list\n2) Client Device list\n>"
        )
        if list_type == "1":
            wb = xlsxwriter.Workbook(f"{all_sites[site_requested]} Unifi Devices.xlsx")
            ws = wb.add_worksheet()
            unifi_devices()
            system("cls")
            break
        elif list_type == "2":
            wb = xlsxwriter.Workbook(f"{all_sites[site_requested]} Client Devices.xlsx")
            ws = wb.add_worksheet()
            client_devices()
            system("cls")
            break
        # Verify input
        else:
            system("cls")
            print("Enter a valid option")
            continue
