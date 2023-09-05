import re

ad_file = "CI Suspected Ads.txt"

with open(ad_file, "r") as read_file:
    lines = read_file.readlines()
ads_list = [line.strip() for line in lines]
# temp_file = "C:\\Users\\darmstrong\\Desktop\\Content Inspection.txt"
temp_file = "Content Inspection.txt"
with open(temp_file, "r") as read_file:
    lines = read_file.readlines()

inspect_li = []
av_li = []
drop_li = []
ad_li = []
for line in lines:
    if "ProxyInspect" in line or "ProxyStrip" in line:
        if "sni=" in line:
            if 'sni=""' in line:
                url = re.search(r'cn="(.+?)"', line)
            else:
                url = re.search(r'sni="(.+?)"', line)
            url = url.group(1)
            if url not in inspect_li:
                inspect_li.append(url)
    elif "ProxyAvScan" in line:
        if "dstname=" in line:
            url = re.search(r'dstname="(.+?)"', line)
            url = url.group(1)
            # if url in ads_list:
            #     ad_li.append(url)
            if url not in inspect_li:
                inspect_li.append(url)
    elif "ProxyDrop" in line:
        if "dstname=" in line:
            url = re.search(r'dstname="(.+?)"', line)
            url = url.group(1)
            # if url in ads_list:
            #     ad_li.append(url)
            if url not in drop_li:
                drop_li.append(url)
for li_item in inspect_li[::-1]:
    for ad in ads_list:
        is_ad = li_item.find(ad)
        if is_ad > 0 and li_item not in ad_li:
            ad_li.append(li_item)
            inspect_li.remove(li_item)

print("\nProxy Inspection:")
for item in inspect_li:
    print(item)
print("\nLikely ad/tracker/analytics:")
for item in ad_li:
    print(item)
ending = input("Press Enter to exit...")
