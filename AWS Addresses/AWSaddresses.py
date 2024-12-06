import json
import os
import time

from json import load as jsonload
from os import system as ossystem
from os import remove as osremove
from time import sleep as timesleep

### PyInstaller command
# pyinstaller --onefile -i"C:\Users\darmstrong\Desktop\pyinstaller\aws.ico"  AWSaddresses.py

# os.system('curl -O https://ip-ranges.amazonaws.com/ip-ranges.json')
# time.sleep(3)
ossystem('curl -O https://ip-ranges.amazonaws.com/ip-ranges.json')
timesleep(3)


### VARIABLES ###
service_types = [
    "AMAZON",
    "S3",
    "EC2",
    "ROUTE53",
    "CLOUDFRONT",
    "GLOBALACCELERATOR",
    "AMAZON_CONNECT",
    "ROUTE53_HEALTHCHECKS_PUBLISHING",
    "CHIME_MEETINGS",
    "CLOUDFRONT_ORIGIN_FACING",
    "CLOUD9",
    "CODEBUILD",
    "API_GATEWAY",
    "ROUTE53_RESOLVER",
    "EBS",
    "EC2_INSTANCE_CONNECT",
    "KINESIS_VIDEO_STREAMS",
    "WORKSPACES_GATEWAYS",
    "AMAZON_APPFLOW",
    "MEDIA_PACKAGE_V2",
    "DYNAMODB"
]

regions = [
    "us-west-1",
    "us-west-2",
    "us-east-1",
    "us-east-2",
    "ca-west-1",
    "GLOBAL",
    "us-gov-east-1",
    "us-gov-west-1",
    "ca-central-1",
    "sa-east-1",
    "eu-west-1",
    "eu-west-2",
    "eu-west-3",
    "eu-central-1",
    "eu-central-2",
    "eu-south-1",
    "eu-south-2",
    "eu-north-1",
    "il-central-1",
    "me-south-1",
    "me-central-1",
    "af-south-1",
    "ap-east-1",
    "ap-south-1",
    "ap-south-2",
    "ap-southeast-1",
    "ap-southeast-2",
    "ap-southeast-3",
    "ap-southeast-4",
    "ap-southeast-5",
    "ap-northeast-1",
    "ap-northeast-2",
    "ap-northeast-3",
]

selected_regions = []
selected_types = []

while True:
    i = 1
    for region in regions:
        print(f"{str(i)}: {region}")
        i += 1
    selected_regions.append(regions[int(input("Select region:\n> "))-1])
    
    ii = 1
    for service_type in service_types:
        print(f"{str(ii)}: {service_type}")
        ii += 1
    selected_types.append(service_types[int(input("Select service type:\n> "))-1])
    end_loop = input("1) Enter another region\n2) Create alias file\n> ")
    if end_loop == "2":
        break
    else:
        continue

num_regions = len(selected_regions)

# Opening JSON file
with open('ip-ranges.json', 'r') as f:
    # data = json.load(f)
    data = jsonload(f)

with open('AWSaddresses.txt', 'w') as txtfile:
    # Using the data as a Python dictionary
    i = 0
    while i < num_regions:
        for entry in data['prefixes']:
            if entry['region'] == f"{selected_regions[i]}" and entry['service'] == f"{selected_types[i]}":
                txtfile.write((f"ipv4,{entry['ip_prefix']}\n"))
        i += 1

# os.remove('ip-ranges.json')
osremove('ip-ranges.json')