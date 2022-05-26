from datetime import date
from img2pdf import convert
from PIL import Image
from os import listdir, remove

def all_dir_paths():
    today = date.today()
    date_today = today.strftime("%Y-%m-%d")
    dir_path = '\\\\FS01\\MSP-SecReview\\weekly\\test'
    dir_list = listdir(dir_path)

    paths = []
    for item in dir_list:
        if item[0].isdigit():
            paths.append(dir_path + "\\" + item + "\\" + date_today)
    return paths

def all_365_paths():
    paths_365 = []
    all_paths = all_dir_paths()
    for item in all_paths:
        client_dir = listdir(item)

        if "o365Mail.png" in client_dir:
            mail_path = (item + "\\" + client_dir[(client_dir.index("o365Mail.png"))])
            paths_365.append(mail_path)
        if "o365Apps.png" in client_dir:
            apps_path = (item + "\\" + client_dir[(client_dir.index("o365Apps.png"))])
            paths_365.append(apps_path)
    return paths_365

def convert_image(in_path):
    pdf_path = (in_path[:-3] + "pdf")
    # opening image
    image = Image.open(in_path)
    # converting into chunks using img2pdf
    pdf_bytes = convert(image.filename)
    # opening or creating pdf file
    file = open(pdf_path, "wb")
    file.write(pdf_bytes)
    image.close()
    file.close()
    remove(in_path)

paths_365 = all_365_paths()
for item in paths_365:
    convert_image(item)