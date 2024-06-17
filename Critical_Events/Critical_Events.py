import configparser
import datetime
import glob
import os
import re
import shutil
import sys
import time
import requests
from datetime import date, timedelta
from os import listdir
import pdfkit
import pyautogui

# import win32com.client
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains

# Path to the directory of your local settings file
SETTINGS_PATH = "C:/Users/svc_scripting/Documents/python_settings"

# path to our wkhtmltopdf executable
PATH_WKHTMLTOPDF = r"C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
CONFIG = pdfkit.configuration(wkhtmltopdf=PATH_WKHTMLTOPDF)

client_list = [
    "0002 - Intermax",
    "0005",
    "0007",
    "0018",
    "0021",
    "0030",
    "0038",
    "0046",
    "0077",
    "0100",
    "0160",
]


def all_dir_paths():
    # Get the subdirectory name based on date
    day_of_week = date.today().weekday()  # Get today's day of the week as an index
    date_monday = date.today() - timedelta(
        days=day_of_week
    )  # Calculate the date of the Monday of this week
    str_date_monday = date_monday.strftime(
        "%Y-%m-%d"
    )  # Convert the complete date of Monday to a string

    dir_path = (
        "\\\\IMX-FS01\\MSP-SecReview\\weekly"  # Upper level directory for all client files
        # "C:\\Users\\darmstrong\\Desktop\\MSP-SecReview\\Weekly"
    )
    dir_list = listdir(dir_path)

    paths = []  # Create a list of all the current week directories
    for item in dir_list:
        if item[0].isdigit():
            paths.append(
                dir_path + "\\" + item + "\\" + str_date_monday
            )  # Add each file to the list as a full filepath
    return paths


def create_ce_html_report(bad_html):
    soup = BeautifulSoup(bad_html, "html.parser")

    html_report = """
    <!doctype html>
    <html>
    <style type="text/css">
        .device-events {
            border-collapse: collapse;
            border-spacing: 0;
            border-color: #aaa;
            //page-break-inside: avoid;
            width: 90%;
        }

        .device-event {
            font-family: Arial, sans-serif;
            font-size: 8px;
            padding: 10px 5px;
            border-style: solid;
            border-width: 0.5px;
            overflow: hidden;
            word-break: normal;
            border-color: #aaa;
            color: #333;
            background-color: lightgray;
            border-color: inherit;
            text-align: left;
            vertical-align: top;
            page-break-inside: avoid;
        }

        .device-event-type {
            font-family: Arial, sans-serif;
            font-size: 9px;
            font-weight: normal;
            padding: 10px 5px;
            border-style: solid;
            border-width: 1px;
            overflow: hidden;
            word-break: normal;
            border-color: #aaa;
            color: #fff;
            background-color: royalblue;
            //page-break-inside: avoid;
        }

        .device-event-id {
            font-weight: 700;
            text-decoration: underline;
            font-size: 8.5px;
        }

        .device-name {
            display: block;
            font-size: .87em;
            margin-top: .33em;
            margin-bottom: .33em;
            padding-left: 5px;
            margin-right: 0;
            font-weight: bold;
        }

        .site {
            display: block;
            font-size: .90em;
            margin-top: 1.67em;
            margin-bottom: .67em;
            margin-left: 2em;
            margin-right: 0;
            font-weight: bold;
        }

        .device {
            margin-bottom: 1.2em;
            padding-left: 2em;
        }
    </style>
    <h2 class="title">Critical Events Report</h2>
    <div class="report">

    """

    site_html = """
    <div class="site">
        <p class="site-name">{}</p>
    """

    device_html = """
    <div class="device">
        <h6 class='device-name'>{}</h6>
    """

    event_type_html = """
    <tr>
        <th class="device-event-type">{}</th>
    </tr>
    """

    event_id_html = """
    <tr>
        <td class="device-event"><span class="device-event-id">{}</span>
    """

    event_description_html = """
        <br>{}</td>
    </tr>
    """
    # get all our text from the report
    # texts = soup.find_all(text=True)

    # now loop trhough and get rid of all the text we don't want
    strings = []
    # our regex patterns we want to match: Sites, devices, EventID and
    # also if it matches Client we will just ignore it. If it doesn't match any of them it is an event description

    client_regex = re.compile("^'Client: ")
    site_regex = re.compile("^'Site: ")
    device_regex = re.compile("^'Device: ")
    event_id_regex = re.compile(r"EventID=\d*")
    event_type_regex = re.compile(r"event\(s\) found")

    # Grab all the stripped strings and loop through them and add them to our html report
    # add them based on regex patterns defined above
    for string in soup.stripped_strings:
        strings.append(repr(string))

    i = 0
    last_item = "start"
    # i used some logic to determine when to add the ending tags and such.
    # it sets whether it is a device/site and checks it with the last item,
    # if the current item is a site and the last item was also a site it will
    # add the ending tags to the site to close off the site tags.
    # if current item is a site and the last item was not a site
    # then we can close off the table and start a new site. same logic for
    # the devices. At the end of each iteration it will set the current to the past and clear the current
    # if last item was a device and the current item is event_type then it will start the table before
    # adding the event type

    while i < len(strings):
        if client_regex.search(strings[i]):
            current_item = "client"
        elif site_regex.search(strings[i]):
            current_item = "site"
        elif device_regex.search(strings[i]):
            current_item = "device"
        elif event_type_regex.search(strings[i]):
            current_item = "event_type"
        elif event_id_regex.search(strings[i]):
            current_item = "event_id"
        else:
            current_item = "event_description"

        if current_item != "event_description" and last_item == "event_id":
            html_report += "</td></tr>"

        if current_item == "site" and last_item == "client":
            html_report += site_html.format(strings[i]).replace("'", "")

        if current_item == "device" and last_item == "site":
            html_report += device_html.format(strings[i]).replace("'", "")

        if current_item == "event_type" and last_item == "device":
            html_report += '<table class="device-events">'
            html_report += event_type_html.format(strings[i]).replace("'", "")

        if current_item == "event_id" and last_item == "event_type":
            html_report += event_id_html.format(strings[i]).replace("'", "")

        if current_item == "event_description" and last_item == "event_id":
            html_report += (
                event_description_html.format(strings[i])
                .replace("'", "")
                .replace(r"\n", " ")
                .replace(r"\t", " ")
            )

        if current_item == "device" and last_item == "event_description":
            html_report += "</table>"
            html_report += "</div>"
            html_report += device_html.format(strings[i]).replace("'", "")

        if current_item == "site" and last_item == "site":
            html_report += "</div>"
            html_report += site_html.format(strings[i]).replace("'", "")

        if current_item == "device" and last_item == "device":
            html_report += "</div>"
            html_report += device_html.format(strings[i]).replace("'", "")

        if current_item == "site" and last_item == "device":
            html_report += "</div>"
            html_report += "</div>"
            html_report += site_html.format(strings[i]).replace("'", "")

        if current_item == "event_id" and last_item == "event_description":
            html_report += event_id_html.format(strings[i]).replace("'", "")

        if current_item == "site" and last_item == "event_description":
            html_report += "</table>"
            html_report += "</div>"
            html_report += "</div>"
            html_report += site_html.format(strings[i]).replace("'", "")

        if current_item == "event_type" and last_item == "event_description":
            html_report += "</table>"
            html_report += '<table class="device-events">'
            html_report += event_type_html.format(strings[i]).replace("'", "")

        # Set our last item to our current item for the next iteration
        last_item = current_item
        # Go to next item
        i += 1

    if last_item == "device":
        html_report += "</div>"
    if last_item == "event_description":
        html_report += "</table> </div>"

    html_report += """
    </div>
    </div>
    </html>
    """

    return html_report


def get_av_reports(driver, dir_date):

    # for Av we want to print in landscape so we need to set our options
    options = {"orientation": "Landscape"}
    # Find all our frames and search for our report bar. once the bar is found in the frame it will click it and break from teh loop
    iframes = driver.find_elements(By.XPATH, "//iframe")
    for f in iframes:
        driver.switch_to.frame(f)
        try:
            report_bar = driver.find_element(By.ID, "maintoolbarreportsbutton")
            if report_bar:
                report_bar.click()
            break
        except:
            driver.switch_to.default_content()
            continue

    time.sleep(1)
    # Find our AV reports item in the menu and click it
    element = driver.find_element(By.ID, "menuitem-1594")
    element.click()
    time.sleep(1)

    # element = driver.find_element(By.CSS_SELECTOR, "menuitem-1594-itemEl")
    # element.click()

    element = driver.find_element(By.ID, "menuitem-1597")
    element.click()

    # Select our second Tab
    clients_window = driver.window_handles[1]
    # Switch to this window
    driver.switch_to.window(clients_window)

    input("Press enter once loaded")
    # Find our protection report button in our new smaller menu and click that
    folders = all_dir_paths()
    client_num = 0
    for folder in folders:
        element = driver.find_element(
            By.ID,
            "clientcombo-1024-inputEl",
        )
        element.click()
        time.sleep(1)
        elements = driver.find_elements(By.CLASS_NAME, "x4-boundlist-item")
        time.sleep(1)
        for e in elements:
            client_id = e.text
            if client_id == client_list[client_num]:
                e.click()
                view_report = driver.find_element(By.ID, "generatebutton-1027")
                view_report.click()
                input("Press enter once loaded")
                # time.sleep(5)
                if client_num == 0:
                    column_select = driver.find_element(
                        By.ID, "viewcolumntoolbarbutton-1059"
                    )
                    column_select.click()
                    time.sleep(1)

                    definition_version = driver.find_element(
                        By.ID, "menucheckitem-1091"
                    )
                    definition_version.click()

                    active_protection = driver.find_element(By.ID, "menucheckitem-1100")
                    active_protection.click()
                    time.sleep(1)
                time.sleep(1)

                # Click the print button
                element = driver.find_element(By.ID, "printbutton-1062-btnIconEl")
                element.click()

                time.sleep(2)
                # switch to the page that now has the print dialogue open
                input("Press enter when ready")
                report_window = driver.window_handles[2]
                driver.switch_to.window(report_window)
                av_html_filename = "AntiVirus.html"
                av_pdf_filename = "MAV.pdf"

                print(f"setup filesystem for {client_id}")
                images_path = f"{folder}\\images"
                if os.path.exists(images_path):
                    print("Folder exists")
                else:
                    os.mkdir(images_path)

                # Click on our element that our client to select our other
                print(f"Generate AV report for {client_id}")

                # grab all of our html
                new_html = driver.find_element(By.TAG_NAME, "html").get_attribute(
                    "outerHTML"
                )

                # initialize our array to store our filenames in
                filenames = []
                filepaths = []
                # get all our images and store the filenames in our list
                images = set(
                    [
                        img.get_attribute("src")
                        for img in driver.find_elements(By.TAG_NAME, "img")
                    ]
                )
                print(f"Download report images: {images}")
                for image in images:
                    filename = image.rsplit("/", 1)[1]
                    print(f"get {image}...", end="")
                    response = requests.get(
                        image,
                        headers={
                            # Pretend to be Chrome 97
                            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36"
                        },
                    )

                    with open(os.path.join(images_path, filename), "wb") as fp_image:
                        fp_image.write(response.content)
                    filepaths.append(f"{images_path}\\{filename}")
                    filenames.append(filename)
                    print(" done.")
                print(filenames)
                print("Switch report images to local paths")
                # now filter through our html and replace our img tags with local paths
                i = 0
                for f in filenames:
                    # get all our img tags
                    new_html.replace(
                        f'img src="/dashboard/images/{f}"', f'<img src="{filepaths[i]}"'
                    )
                    i += 1

                # print("remove print button")
                # # get rid of that ugly print button
                # new_html = re.sub(
                #     r"<div id=\"printButton\" class=\"print-button-holder hidden-print\"><button>Print</button></div>",
                #     "",
                #     new_html,
                # )

                print("Save HTML to file")
                f = open(av_html_filename, "w")
                f.write(new_html)
                # close our file
                f.close()

                print("Convert HTML to PDF")
                # this is in a try/except as there was an issue with a couple external styling links that wktohtml has open issues
                # due to being unable to handle or suppress these errors, the report still comes out looking fine though
                try:
                    # convert it to a pdf
                    pdfkit.from_file(
                        av_html_filename,
                        av_pdf_filename,
                        configuration=CONFIG,
                        options=options,
                    )
                except Exception as e:
                    print(e)

                print("Clean up")
                # remove our html file
                os.remove(av_html_filename)

                # remove our image folder now that we have our pdf
                shutil.rmtree(images_path)

                # close our page
                driver.close()

                # Switch back to our
                window = driver.window_handles[1]
                # Switch to this window
                driver.switch_to.window(window)

                element = driver.find_element(By.ID, "ext4-ext-gen1142")
                # Click to open our drop down menu
                element.click()
        # close our AV report page now that we have all of our reports
        driver.close()


def get_critical_events(driver, dir_date):

    # for Critical events we want to print in portrait so we need to set our options
    options = {"orientation": "portrait"}
    iframes = driver.find_elements(By.XPATH, "//iframe")
    for f in iframes:
        driver.switch_to.frame(f)
        try:
            report_bar = driver.find_element(By.ID, "maintoolbarreportsbutton")
            if report_bar:
                report_bar.click()
            break
        except:
            driver.switch_to.default_content()
            continue

    time.sleep(3)
    # Find our Critical Events report and click it
    element = driver.find_element(
        # By.CSS_SELECTOR, "[data-ui-comp-name='rm-critical-event']"
        By.ID,
        "menuitem-1561-itemEl",
    )
    element.click()
    time.sleep(3)
    # Select our second Tab
    window = driver.window_handles[1]
    # Switch to this window
    driver.switch_to.window(window)

    # Get our element with our client list
    element = driver.find_element(By.CSS_SELECTOR, "[role='combobox']")
    # CLick to open our drop down menu
    element.click()
    time.sleep(1)
    # find our element that contains our list of items
    elements = driver.find_elements(By.CLASS_NAME, "x4-boundlist-item")
    folders = all_dir_paths()
    client_num = 0
    for folder in folders:
        for e in elements:
            if e.text == client_list[client_num]:
                e.click()
                view_report = driver.find_element(By.ID, "generatebutton-1024")
                view_report.click()
                time.sleep(5)

                event_html_filename = "C:\\tempdir\\CriticalEvents.html"
                event_pdf_filename = folder + "\\Critical Events.pdf"

                # now we must get our html in order to print it
                element = driver.find_element(
                    By.ID, "criticaleventsreportgrid-1031-outerCt"
                )
                event_html = element.get_attribute("outerHTML")
                f = open(event_html_filename, "w", errors="ignore")
                f.write("<!doctype html>\n<html>\n")
                # we need to replace the string \u200e as python doesn't know how to decode this string. and it is just supposed to be blank
                f.write(event_html)
                f.write("\n</html>")
                # close our file
                f.close()

                # now we need to parse our ugly html from Solarwinds and get the data to put into my HTML report I made
                # note I did this due to them using <div display: table> rather than actually using Table tags which
                # causes issues when converting to wkhtmltopdf, basically it doesn't know how to handle line breaks
                # and makes giant yellow boxes if the event box goes onto the next page
                bad_html = open(event_html_filename, "r")
                new_report_html = create_ce_html_report(bad_html)
                bad_html.close()

                f = open(event_html_filename, "w")
                f.write(new_report_html)
                # close our file
                f.close()

                # convert it to a pdf
                pdfkit.from_file(
                    event_html_filename,
                    event_pdf_filename,
                    configuration=CONFIG,
                    options=options,
                )

                # remove our html file
                os.remove(event_html_filename)

                # re-open our drop down menu
                element = driver.find_element(By.CSS_SELECTOR, "[role='combobox']")

                # Click to open our drop down menu
                element.click()
                client_num += 1
                time.sleep(1)
                break
    # close our page once we're done
    driver.close()


def main():
    config = configparser.ConfigParser()
    with open("settings.ini", "r") as fp_config: 
        config.read_file(fp_config)

    # get our date for our directories to store our files in
    dir_date = str(datetime.date.today() - datetime.timedelta(days=0))

    chrome_options = Options()

    # prefs = {"download.default_directory": r"//fs01/MSP-SecReview/weekly"}
    # chrome_options.add_experimental_option("prefs", prefs)

    # add arg to make this run as headless
    # chrome_options.add_argument('--headless')
    driver = webdriver.Chrome(options=chrome_options)
    driver.get("https://dashboard.systemmonitor.us/")

    element = driver.find_element(By.ID, "email-field")
    element.send_keys(config.get("Solarwinds", "username"))
    element = driver.find_element(By.ID, "email-submit")
    element.click()
    time.sleep(1)

    element = driver.find_element(By.ID, "password-field")
    element.send_keys(config.get("Solarwinds", "password"))
    element = driver.find_element(By.ID, "password-submit")
    element.click()
    time.sleep(1)
    code = int(input("Enter OTP:"))
    element = driver.find_element(By.ID, "code-field")
    element.send_keys(code)
    # element.send_keys(config.get("Solarwinds", "code"))
    element = driver.find_element(By.ID, "verify-submit")
    element.click()

    # # See if there is an ad pop-up and click them until they are all gone
    # while True:
    #     try:
    #         # element = driver.find_element_by_class_name("_pendo-close-guide")
    #         element = driver.find_element(By.CLASS_NAME, "_pendo-close-guide")
    #         element.click()
    #     except:
    #         pass
    #         break

    # driver.implicitly_wait(20)

    print("********** CRITICAL EVENTS REPORTS **********")
    time.sleep(60)
    get_critical_events(driver, dir_date)

    # # Grab our main page and switch back to it
    # window = driver.window_handles[0]
    # driver.switch_to.window(window)

    # print("********* MANAGED ANTIVIRUS REPORTS *********")
    # get_av_reports(driver, dir_date)

    # Grab our main page and switch back to it and close it
    window = driver.window_handles[0]
    driver.switch_to.window(window)
    driver.close()

    # remove our driver once we're done here
    driver.quit()


if __name__ == "__main__":
    main()
