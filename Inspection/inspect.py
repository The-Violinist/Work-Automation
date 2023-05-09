import re

temp_file = "C:\\Users\\darmstrong\\Desktop\\Content Inspection.txt"
with open(temp_file, "r") as read_file:
    lines = read_file.readlines()

inspect_li = []
av_li = []
for line in lines:
    if "ProxyInspect" in line or "ProxyStrip" in line:
        if "sni=" in line:
            if 'sni=""' in line:
                url = re.search(r'cn="(.+?)"', line)
            else:
                url = re.search(r'sni="(.+?)"', line)
            if url.group(1) not in inspect_li:
                inspect_li.append(url.group(1))
    elif "ProxyAvScan" in line:
        if "dstname=" in line:
            url = re.search(r'dstname="(.+?)"', line)
            if url.group(1) not in av_li:
                av_li.append(url.group(1))

print("\nProxy Inspection:")
for item in inspect_li:
    print(item)
print("\nAV Scan:")
for item in av_li:
    print(item)
ending = input("Press Enter to exit...")
