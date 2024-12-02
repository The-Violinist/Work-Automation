import json

# Opening JSON file
with open('ip-ranges.json', 'r') as f:
    data = json.load(f)

with open('AWSaddresses.txt', 'w') as txtfile:
    # Using the data as a Python dictionary
    for entry in data['prefixes']:
        if entry['region'] == "us-east-1" and entry['service'] == "AMAZON":
            txtfile.write((f"ipv4,{entry['ip_prefix']}\n"))
        elif entry['region'] == "us-east-2" and entry['service'] == "AMAZON":
            txtfile.write((f"ipv4,{entry['ip_prefix']}\n"))
        elif entry['region'] == "us-west-1" and entry['service'] == "AMAZON":
            txtfile.write((f"ipv4,{entry['ip_prefix']}\n"))
        elif entry['region'] == "us-west-2" and entry['service'] == "AMAZON":
            txtfile.write((f"ipv4,{entry['ip_prefix']}\n"))
        elif entry['region'] == "us-west-2" and entry['service'] == "S3":
            txtfile.write((f"ipv4,{entry['ip_prefix']}\n"))
        elif entry['region'] == "GLOBAL" and entry['service'] == "AMAZON":
            txtfile.write((f"ipv4,{entry['ip_prefix']}\n"))

