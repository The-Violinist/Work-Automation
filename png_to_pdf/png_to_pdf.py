from datetime import date, timedelta
from img2pdf import convert
from PIL import Image
from os import listdir, remove

# Create a list of all the weekly folder paths
def all_dir_paths():
    # Get the subdirectory name based on date
    day_of_week = date.today().weekday()
    date_monday = date.today() - timedelta(days=day_of_week)
    str_date_monday = date_monday.strftime("%Y-%m-%d")
    # Parent for all client directories
    dir_path = "\\\\FS01\\MSP-SecReview\\weekly"
    dir_list = listdir(dir_path)
    # Create a list of all the current week directories
    paths = []
    for item in dir_list:
        if item[0].isdigit():
            paths.append(dir_path + "\\" + item + "\\" + str_date_monday)
    return paths


# Create a list of paths to all the png files
def all_365_paths():
    paths_365 = []
    all_paths = all_dir_paths()
    for item in all_paths:
        client_dir = listdir(item)
        # Find all png files and add file paths to paths_365
        if "o365Mail.png" in client_dir:
            mail_path = item + "\\" + client_dir[(client_dir.index("o365Mail.png"))]
            paths_365.append(mail_path)
        if "o365Apps.png" in client_dir:
            apps_path = item + "\\" + client_dir[(client_dir.index("o365Apps.png"))]
            paths_365.append(apps_path)
    return paths_365


# Convert the png files to pdf
def convert_image(in_path):
    # Create name for output file
    pdf_path = in_path[:-3] + "pdf"
    # Open image
    image = Image.open(in_path)
    # Convert into chunks using img2pdf
    pdf_bytes = convert(image.filename)
    # Create pdf file
    file = open(pdf_path, "wb")
    file.write(pdf_bytes)
    image.close()
    file.close()
    remove(in_path)


paths_365 = all_365_paths()
for item in paths_365:
    convert_image(item)

# pyinstaller.exe --onefile png_to_pdf.py
# spec file pathex: "C:\Python310\Lib\site-packages"
# pyinstaller.exe --onefile -i"pdf.ico" png_to_pdf.spec
