import configparser
import datetime
import glob
import json
import os
import random
import shutil
import subprocess
import sys
import time
import tempfile
import codecs

import requests

import pdfkit
from lxml import etree, html

# Ignore decoding errors
codecs.register_error("strict", codecs.ignore_errors)

# list of all client ids (These are the IDs that Solarwinds generate for them when the client is created on the Dashboard)
# If you add anyone to this list it will pull the reports for them as well
CLIENT_IDS = ['0002 - Intermax', '0005', '0007', '0018', '0021', '0030', '0038', '0046', '0077', '0100', '0160', '0170', '0171']
# CLIENT_IDS = ['0005']

# Url to log into Solarwinds RMM
# LOGIN_URL = "https://sso.navigatorlogin.com/Account/Login?ReturnUrl=%2Fconnect%2Fauthorize%2Fcallback%3Fclient_id%3D5e332073-1394-4330-98c1-92b4d2e78c09%26response_type%3Dcode%26scope%3Dopenid%2520email%2520profile%26state%3D%257B%2522data%2522%253Anull,%2522state%2522%253A%2522d712eb63c62f18b1a37381adbcae15ba%2522,%2522originUrl%2522%253A%2522https%253A%255C%252F%255C%252Fdashboard.systemmonitor.us%255C%252Fmsp_sso.php%2522%257D%26redirect_uri%3Dhttps%253A%252F%252Fwww.systemmonitor.us%252Fdashboard%252Fmsp_sso.php"
LOGIN_URL = "https://sso.navigatorlogin.com/Account/Login?ReturnUrl=%2Fconnect%2Fauthorize%2Fcallback%3Fclient_id%3Dfa1c6d98-0795-4396-a172-78a611b8ee93%26response_type%3Dcode%26scope%3Dopenid%2520email%2520profile%2520offline_access%26state%3D%257B%2522data%2522%253Anull,%2522state%2522%253A%2522d13927236c10457e5dc1697301525738%2522,%2522originUrl%2522%253A%2522https%253A%255C%252F%255C%252Fdashboard.systemmonitor.us%255C%252Fmsp_sso.php%2522%257D%26redirect_uri%3Dhttps%253A%252F%252Fwww.systemmonitor.us%252Fdashboard%252Fmsp_sso.php"

# USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.87 Safari/537.36"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
PATH_WKHTMLTOPDF = r'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'
CONFIG = pdfkit.configuration(wkhtmltopdf=PATH_WKHTMLTOPDF)

#AUTH_CODE = ''

def get_sw_client_ids(token):

    # initialize our dictionary
    sw_ids = {}

    # this guy makes an api call to Solarwinds which returns a list of the client IDs that they generate
    sw_api_url = f"https://www.systemmonitor.us/api/?apikey={token}&service=list_clients"
    client_result = requests.get(sw_api_url)
    xml = etree.fromstring(client_result.content)
    xml_clients = xml.findall(".//client")
    for client in xml_clients:
        # Get the Value of the name attribute which holds the Id as we know and love them at intermax
        imax_id = client.find('name').text
        # Get the value of the clientid tag which is the id that we will use in our report url
        sw_id = client.find('clientid').text
        sw_ids[imax_id] = sw_id
    return sw_ids


def get_take_control(sw_cid, cid, start, end, cookie, session, dir_date):

    # We sleep for 3 - 7 seconds as this is about how long it takes me to manually click on the reports
    # and we don't want to make SolarWinds angry by querying their server too quickly and I made it
    # random as to better impersonate that a human is doing this
    # print(f"Getting Take Control Report for {cid}")
    # time.sleep(2)

    # name/path of file take contorl html report
    tc_html_filename = os.path.join("C:", os.sep, "tempdir", "Patches", f"{cid}-TakeControlReport.html")
    tc_pdf_filename = os.path.join("C:", os.sep, "tempdir", "Patches", f"{cid}-TakeControlReport.pdf")
    # client_path = glob.glob(os.path.join("C:\\Users\\darmstrong\\Desktop\\MSP-SecReview\\Weekly", f"{cid}*")).pop()
    client_path = glob.glob(os.path.join("\\\\imx-fs01", "MSP-SecReview", "weekly", f"{cid}*")).pop()
    tc_destination = os.path.join(client_path, dir_date, "TakeControlReport.pdf")
    print(tc_destination)

    # TODO: These can be one single format step, change to Requests URL builder
    # path to report used in header, format with client name and start/end dates
    tc_path = f"data_processor.php?function=take_control_report&clientid={sw_cid}&startdate={start}T00%3A00%3A00&enddate={end}T00%3A00%3A00"

    # format our report_url with our path since they need to match
    tc_report_url = f"https://dashboard.systemmonitor.us/{tc_path}"
    print(tc_report_url)
    '''
    # Got these headers from Chrome Dev Tools
    # TODO: Try removing everything except Cookie and user-agent
    tc_headers = {'authority': 'dashboard.systemmonitor.us',
                    'method': 'GET',
                    'path': tc_path,
                    'scheme': 'https',
                    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
                    'accept-encoding': 'gzip, deflate, br',
                    'accept-language': 'en-US, en;q = 0.9',
                    'cache-control': 'max-age = 0',
                    'cookie': cookie,
                    'upgrade-insecure-requests': '1',
                    'user-agent': USER_AGENT
                    }

    # Scrape url for the report
    result = session.get(
        tc_report_url, headers=tc_headers)
    # Need to decode the result as it is returned as bytes rather than a string
    # TODO: find out how often these switch 
    print(f'Result code: {result}')
    try:
        tc_decoded_content = result.content.decode("utf-8")
    except:
        tc_decoded_content = result.content.decode("windows-1252")
    # open our file and Write our decoded content to it
    f = open(tc_html_filename, 'w')
    f.write(tc_decoded_content)
    # close our file
    f.close()

    try:
        pdfkit.from_file(tc_html_filename, tc_pdf_filename, configuration=CONFIG)
    except Exception as e:
        print(e)

    shutil.copyfile(tc_pdf_filename, tc_destination)

    print("Remove temporary files.")
    os.remove(tc_html_filename)
    os.remove(tc_pdf_filename)
    '''
def get_patch_management(sw_cid, cid, cookie, session, dir_date):
    # print(f"Getting Patch Management Overview report for {cid}")
    # time.sleep(2)

    patches_html_filename = os.path.join("C:", os.sep, "tempdir", "Patches", f"{cid}-PatchManagement.html")
    patches_pdf_filename = os.path.join("C:", os.sep, "tempdir", "Patches", f"{cid}-PatchManagement.pdf")

    # client_path = ("C:\\Users\\darmstrong\\Desktop\\MSP-SecReview\\Weekly")
    client_path = glob.glob(os.path.join("\\\\imx-fs01", "MSP-SecReview", "weekly", f"{cid}*")).pop()
    patches_destination = os.path.join(client_path, dir_date, "PatchManagement.pdf")
    # patches path
    patches_path = "/data_processor_classlib.php?function=patch/patch_overview_report&action=generateReport&client={}&patch_status=".format(
        sw_cid) + "{%221%22:true,%222%22:true,%224%22:false,%228%22:false,%2216%22:true,%2232%22:false,%2264%22:false,%22128%22:false}&render_by=2&format=html"

    # do the same with Patch Management now
    patches_report_url = 'https://dashboard.systemmonitor.us{}'.format(
        patches_path)
    print(patches_report_url)
    print(patches_destination)

    '''
    # TODO: Try wiht just cookies and user-agent
    patches_headers = {'authority': 'dashboard.systemmonitor.us',
                        'method': 'GET',
                        'path': patches_path,
                        'scheme': 'https',
                        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
                        'accept-encoding': 'gzip, deflate, br',
                        'accept-language': 'en-US, en;q = 0.9',
                        'cache-control': 'max-age = 0',
                        'cookie': cookie,
                        'upgrade-insecure-requests': '1',
                        'user-agent': USER_AGENT
                        }
    # Scrape url for the report
    result = session.get(patches_report_url, headers=patches_headers)
    # Need to decode the result as it is returned as bytes rather than a string
    patches_decoded_content = result.content.decode("utf-8")
    # open our file and Write our decoded content to it
    f = open(patches_html_filename, 'w')
    f.write(patches_decoded_content)
    # close our file
    f.close()

    try:
        pdfkit.from_file(patches_html_filename, patches_pdf_filename, configuration=CONFIG)
    except Exception as e:
        print(e)

    shutil.copyfile(patches_pdf_filename, patches_destination)

    print("Remove temporary files.")
    os.remove(patches_html_filename)
    os.remove(patches_pdf_filename)
    '''
def get_web_protection(sw_cid, cid, start, end, cookie, session, dir_date):
    
    # print(
    #     "Getting Web Protection report for {}, {} seconds remaining".format(cid))
    # time.sleep(2)

    wp_html_filename = os.path.join("C:", os.sep, "tempdir", "Patches", f"{cid}-WebProtection.html")
    wp_pdf_filename = os.path.join("C:", os.sep, "tempdir", "Patches", f"{cid}-WebProtection.pdf")

    # client_path = ("C:\\Users\\darmstrong\\Desktop\\MSP-SecReview\\Weekly")
    client_path = glob.glob(os.path.join("\\\\imx-fs01", "MSP-SecReview", "weekly", f"{cid}*")).pop()
    wp_destination = os.path.join(client_path, dir_date, "WebProtection.pdf")
    # wp path
    wp_path = "/data_processor_classlib.php?function=webprotection/web_protection_overview_report&action=generateReport&varid=44541&clientid={}&siteid=0&deviceid=0&startdate={}%2000:00:00&enddate={}%2000:00:00&summary=true&security=true&filtering=true&bandwidth=true".format(sw_cid, start, end)

    # url for the report
    wp_report_url = 'https://dashboard.systemmonitor.us{}'.format(
        wp_path)
    print(wp_report_url)
    print(wp_destination)

    '''
    # TODO: Try to use just cookies and user-agent
    wp_headers = {'authority': 'dashboard.systemmonitor.us',
                        'method': 'GET',
                        'path': wp_path,
                        'scheme': 'https',
                        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
                        'accept-encoding': 'gzip, deflate, br',
                        'accept-language': 'en-US, en;q = 0.9',
                        'cache-control': 'max-age = 0',
                        'cookie': cookie,
                        'upgrade-insecure-requests': '1',
                        'user-agent': USER_AGENT
                        }
    # Scrape url for the report
    result = session.get(
        wp_report_url, headers=wp_headers)
    # Need to decode the result as it is returned as bytes rather than a string
    wp_decoded_content = result.content.decode("utf-8")
    # open our file and Write our decoded content to it
    f = open(wp_html_filename, 'w')
    f.write(wp_decoded_content)
    # close our file
    f.close()

    try:
        pdfkit.from_file(wp_html_filename, wp_pdf_filename, configuration=CONFIG)
    except Exception as e:
        print(e)

    shutil.copyfile(wp_pdf_filename, wp_destination)

    print("Remove temporary files.")
    os.remove(wp_html_filename)
    os.remove(wp_pdf_filename)
    '''
def main():
    config = configparser.ConfigParser()
    with open("settings.ini", "r") as fp_config:
        config.read_file(fp_config)

    # Get our Start and end times to run this on, A week ago and yesterday and store them as strings
    start_date = str(datetime.date.today() - datetime.timedelta(days=7))
    end_date = str(datetime.date.today() - datetime.timedelta(days=1))
    dir_date = str(datetime.date.today() - datetime.timedelta(days=0))

    # Get our Dictionary of IDs
    ids = get_sw_client_ids(config.get("Solarwinds", "api_token"))

    # Open our requests session
    session_requests = requests.session()

    # Set our User Agent to the one we took from Chrome
    session_requests.headers['User-Agent'] = USER_AGENT


    # Get our Request verification Token from the login page, this is required when passing our credentials
    result = session_requests.get(LOGIN_URL)
    tree = html.fromstring(result.text)
    verification_token = list(
        set(tree.xpath("//input[@name='__RequestVerificationToken']/@value")))

    # Create payload which is used to authenticate to the site
    # I got the names of these fields from Chrome Dev Tools
    payload = {
        "Email": config.get("Solarwinds", "username"),
        "Password": config.get("Solarwinds", "password"),
        "__RequestVerificationToken": verification_token,
        "Code": config.get("Solarwinds", "code"),
        # "email-field": config.get("Solarwinds", "code")
    }
    # Perform login
    result = session_requests.post(
        LOGIN_URL, data=payload, headers=dict(referer=LOGIN_URL))

    # Grab our cookies
    result_cookies = result.request.headers['cookie']

    password_url = "https://sso.navigatorlogin.com/Account/LoginEmail?returnurl=%2Fconnect%2Fauthorize%2Fcallback%3Fclient_id%3Dfa1c6d98-0795-4396-a172-78a611b8ee93%26response_type%3Dcode%26scope%3Dopenid%2520email%2520profile%2520offline_access%26state%3D%257B%2522data%2522%253Anull,%2522state%2522%253A%2522d2b4daf5556a1cbc1615e945efd623f4%2522,%2522originUrl%2522%253A%2522https%253A%255C%252F%255C%252Fdashboard.systemmonitor.us%255C%252Fmsp_sso.php%2522%257D%26redirect_uri%3Dhttps%253A%252F%252Fwww.systemmonitor.us%252Fdashboard%252Fmsp_sso.php"
    # Perform login
    result = session_requests.post(
        password_url, data=payload, headers=dict(referer=password_url))

    result_cookies = result.request.headers['cookie']

    # The url that we use to enter out 2 Factor Authentication Code
    verify_url = 'https://sso.navigatorlogin.com/Account/VerifyCode'
    # verify_url = 'https://sso.navigatorlogin.com/Account/VerifyCode?ReturnUrl=%2Fconnect%2Fauthorize%2Fcallback%3Fclient_id%3Dfa1c6d98-0795-4396-a172-78a611b8ee93%26response_type%3Dcode%26scope%3Dopenid%2520email%2520profile%2520offline_access%26state%3D%257B%2522data%2522%253Anull,%2522state%2522%253A%252296fffa41c5ae77801ce6ce9c2ff308e9%2522,%2522originUrl%2522%253A%2522https%253A%255C%252F%255C%252Fdashboard.systemmonitor.us%255C%252Fmsp_sso.php%2522%257D%26redirect_uri%3Dhttps%253A%252F%252Fwww.systemmonitor.us%252Fdashboard%252Fmsp_sso.php&RememberMe=False&Provider=Authenticator'
    
    # Grab our headers from earlier
    verify_headers = result.request.headers

    # Form data to enter our 2FA code
    verify_data = {
        'Provider': 'Authenticator',
        'RememberMe': 'False',
        # 'ReturnUrl': "/connect/authorize/callback?client_id=5e332073-1394-4330-98c1-92b4d2e78c09&response_type=code&scope=openid%20email%20profile&state=%7B%22data%22%3Anull,%22state%22%3A%225c7743748b4f7ca094ebec1a91d71a70%22,%22originUrl%22%3A%22https%3A%5C%2F%5C%2Fdashboard.systemmonitor.us%5C%2Fmsp_sso.php%22%7D&redirect_uri=https%3A%2F%2Fwww.systemmonitor.us%2Fdashboard%2Fmsp_sso.php",
        'ReturnUrl': "/connect/authorize/callback?client_id=fa1c6d98-0795-4396-a172-78a611b8ee93&amp;response_type=code&amp;scope=openid%20email%20profile%20offline_access&amp;state=%7B%22data%22%3Anull,%22state%22%3A%2296fffa41c5ae77801ce6ce9c2ff308e9%22,%22originUrl%22%3A%22https%3A%5C%2F%5C%2Fdashboard.systemmonitor.us%5C%2Fmsp_sso.php%22%7D&amp;redirect_uri=https%3A%2F%2Fwww.systemmonitor.us%2Fdashboard%2Fmsp_sso.php",
        # 'Code': "885633",
        '__RequestVerificationToken': verification_token,
        'RememberBrowser': 'false'
    }

    # Make the request and we will use our cookie from this in order to use our authenticated session for our report calls
    verify_result = session_requests.post(verify_url, data=verify_data, headers=verify_headers)
    verify_cookies = result.request.headers['cookie']
    
    for c in CLIENT_IDS:

        # Change our cwd so we can use glob to find sub directories
        # os.chdir('//fs01/MSP-SecReview/weekly')

        # Get our sw client ID using the imax client id
        sw_client_id = ids[c]
        
        # get_take_control(sw_client_id, c, start_date, end_date, result_cookies, session_requests, dir_date)
        
        # get_patch_management(sw_client_id, c, result_cookies, session_requests, dir_date)

        # get_web_protection(sw_client_id, c, start_date, end_date, result_cookies, session_requests, dir_date)
if __name__ == '__main__':
    main()
