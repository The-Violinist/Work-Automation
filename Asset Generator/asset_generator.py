print('''
                                                                               ██╗    ██╗███████╗██╗      ██████╗ ██████╗ ███╗   ███╗███████╗    ████████╗ ██████╗                                                                  
                                                                               ██║    ██║██╔════╝██║     ██╔════╝██╔═══██╗████╗ ████║██╔════╝    ╚══██╔══╝██╔═══██╗                                                                 
                                                                               ██║ █╗ ██║█████╗  ██║     ██║     ██║   ██║██╔████╔██║█████╗         ██║   ██║   ██║                                                                 
                                                                               ██║███╗██║██╔══╝  ██║     ██║     ██║   ██║██║╚██╔╝██║██╔══╝         ██║   ██║   ██║                                                                 
                                                                               ╚███╔███╔╝███████╗███████╗╚██████╗╚██████╔╝██║ ╚═╝ ██║███████╗       ██║   ╚██████╔╝                                                                 
                                                                               ╚══╝╚══╝ ╚══════╝╚══════╝ ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚══════╝       ╚═╝    ╚═════╝                                                                  
                                                 ████████╗██╗  ██╗███████╗    ██████╗  █████╗ ███████╗██╗  ██╗██████╗  ██████╗  █████╗ ██████╗ ██████╗      █████╗ ███████╗███████╗███████╗████████╗                          
                                                 ╚══██╔══╝██║  ██║██╔════╝    ██╔══██╗██╔══██╗██╔════╝██║  ██║██╔══██╗██╔═══██╗██╔══██╗██╔══██╗██╔══██╗    ██╔══██╗██╔════╝██╔════╝██╔════╝╚══██╔══╝                          
                                                    ██║   ███████║█████╗      ██║  ██║███████║███████╗███████║██████╔╝██║   ██║███████║██████╔╝██║  ██║    ███████║███████╗███████╗█████╗     ██║                             
                                                    ██║   ██╔══██║██╔══╝      ██║  ██║██╔══██║╚════██║██╔══██║██╔══██╗██║   ██║██╔══██║██╔══██╗██║  ██║    ██╔══██║╚════██║╚════██║██╔══╝     ██║                             
                                                    ██║   ██║  ██║███████╗    ██████╔╝██║  ██║███████║██║  ██║██████╔╝╚██████╔╝██║  ██║██║  ██║██████╔╝    ██║  ██║███████║███████║███████╗   ██║                             
                                                    ╚═╝   ╚═╝  ╚═╝╚══════╝    ╚═════╝ ╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝╚═════╝  ╚═════╝ ╚═╝  ╚═╝╚═╝  ╚═╝╚═════╝     ╚═╝  ╚═╝╚══════╝╚══════╝╚══════╝   ╚═╝                             
                                    ███████╗██████╗ ██████╗ ███████╗ █████╗ ██████╗ ███████╗██╗  ██╗███████╗███████╗████████╗     ██████╗ ███████╗███╗   ██╗███████╗██████╗  █████╗ ████████╗ ██████╗ ██████╗ ██╗
                                    ██╔════╝██╔══██╗██╔══██╗██╔════╝██╔══██╗██╔══██╗██╔════╝██║  ██║██╔════╝██╔════╝╚══██╔══╝    ██╔════╝ ██╔════╝████╗  ██║██╔════╝██╔══██╗██╔══██╗╚══██╔══╝██╔═══██╗██╔══██╗██║
                                    ███████╗██████╔╝██████╔╝█████╗  ███████║██║  ██║███████╗███████║█████╗  █████╗     ██║       ██║  ███╗█████╗  ██╔██╗ ██║█████╗  ██████╔╝███████║   ██║   ██║   ██║██████╔╝██║
                                    ╚════██║██╔═══╝ ██╔══██╗██╔══╝  ██╔══██║██║  ██║╚════██║██╔══██║██╔══╝  ██╔══╝     ██║       ██║   ██║██╔══╝  ██║╚██╗██║██╔══╝  ██╔══██╗██╔══██║   ██║   ██║   ██║██╔══██╗╚═╝
                                    ███████║██║     ██║  ██║███████╗██║  ██║██████╔╝███████║██║  ██║███████╗███████╗   ██║       ╚██████╔╝███████╗██║ ╚████║███████╗██║  ██║██║  ██║   ██║   ╚██████╔╝██║  ██║██╗
                                    ╚══════╝╚═╝     ╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝╚═════╝ ╚══════╝╚═╝  ╚═╝╚══════╝╚══════╝   ╚═╝        ╚═════╝ ╚══════╝╚═╝  ╚═══╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝╚═╝
    ''')
import os
import xml.etree.ElementTree as ET
import csv

# Get input file from user to create XML tree
while True:
    file = input("Enter XML filename or press enter to exit:\n>")
    if file == "":
        os.system('cls')
        quit()
    try:
        tree = ET.parse(file)
        root = tree.getroot()
        break
    except:
        print("That file does not exist.\n\n")
        continue

###FUNCTIONS###
#Get machine name from user input. Send to machine_loop as name

def get_machine_name():
    machines = []
    for child in root:
        for child1 in child:
            for child2 in child1.findall("name"):
                machines.append(child2.text)
    i = 1
    print("Select a computer from the list, or press enter to quit:\n")
    for item in machines:
        print(f"{i}) {item}")
        i += 1
    selection = (input("\n>"))
    if selection == "":
        os.system('cls')
        quit()
    else:
        try:
            #selection = int(selection)
            return (machines[int(selection)-1])
        except:
            os.system('cls')
            quit()
    #return (machines[selection-1])

# Loop through menu options
def machine_loop():
    while True:
        ask = input("Enter option number:\n\n1) Create software spreadsheet\n2) Create hardware spreadsheet\n3) Exit program\n\n\n>")
        if ask == "1":
            # Pass to get_software as machine_name
            name = get_machine_name()
            get_software(name)
            one_more = input("Would you like to perform another action?\n>")
            if one_more == "YES" or one_more == "yes" or one_more == "Y" or one_more == "y":
                continue
            else:
                break
        elif ask == "2":
            name = get_machine_name()
            get_hardware(name)
            one_more = input("Would you like to perform another action?\n>")
            if one_more == "YES" or one_more == "yes" or one_more == "Y" or one_more == "y":
                continue
            else:
                break
        elif ask == "3":
            break
        else:
            print("Please enter a valid selection.\n\n")

# Write headers for software
def sft_head(headlist, mn):
    with open(f'{mn}_software.csv', 'w') as file:
        dw = csv.DictWriter(file, delimiter=',', fieldnames=headlist)
        dw.writeheader()

# Write headers for hardware
def hdw_head(headlist, mn):
    with open(f'{mn}_hardware.csv', 'w') as file:
        dw = csv.DictWriter(file, delimiter=',', fieldnames=headlist)
        dw.writeheader()

# Create csv writer for software
def write_to_sft_csv(app, mn):
    with open(f'{mn}_software.csv', 'a', newline='') as f:
        writer = csv.writer(f)
        for item in app:
            writer.writerows([item])

# Create csv writer for hardware
def write_to_hdw_csv(app, mn):
    with open(f'{mn}_hardware.csv', 'a', newline='') as f:
        writer = csv.writer(f)
        for item in app:
            writer.writerows([item])

# Extract software information from XML file
def get_software(machine_name):
    for child in root:
        for child1 in child:
            for child2 in child1.findall("name"):
                if child2.text == machine_name:
                    for child2 in child1.findall("software"):
                        headers = ['Software','version']
                        sft_head(headers, machine_name)
                        for child3 in child2:
                            arr = []
                            arr1 = []
                            for child4 in child3.findall("name"):
                                application = child4.text
                                arr.append(application)
                            for child4 in child3.findall("version"):
                                version = child4.text
                                arr.append(version)
                            arr1.append(arr)
                            write_to_sft_csv(arr1, machine_name)

# Extract hardware information from XML file
def get_hardware(machine_name):
    for child in root:
        for child1 in child:
            for child2 in child1.findall("name"):
                if child2.text == machine_name:
                    for child2 in child1.findall("hardware"):
                        headers = ['Type','Name','Manufacturer']
                        hdw_head(headers, machine_name)
                        for child3 in child2:
                            arr = []
                            arr1 = []
                            for child4 in child3.findall("type"):
                                hdw_type = child4.text
                                arr.append(hdw_type)
                            for child4 in child3.findall("name"):
                                hdw_name = child4.text
                                arr.append(hdw_name)
                            for child4 in child3.findall("manufacturer"):
                                manufacturer = child4.text
                                arr.append(manufacturer)                            
                                arr1.append(arr)
                            write_to_hdw_csv(arr1, machine_name)

###MAIN###
machine_loop()
os.system('cls')
###END###