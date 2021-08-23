# Simple python script to unlock protected Excel files.
# Each Excel file is a hidden .ZIP archive containing .XML files for each worksheet.
# Within these .XML is a <SheetProtection> tag locked sheet and a <workbookProtection> tag
# containing the hashed passwords etc.
# Retrieving and cracking these passwords are of course possible, but by simply removing the
# <SheetProtection> / <workbookProtection> tags, the workbook will become completely unlocked.

import os
import sys
import zipfile
from codecs import open
from re import search
from shutil import copyfile, rmtree, make_archive
from tkinter import filedialog, Tk

# Initialize and close Tkinter root window
root = Tk()
root.withdraw()

# Open Tkinter and choose file from from current directory.
open_file = filedialog.askopenfilename(initialdir=os.getcwd())

if not open_file:
    exit()

# set file variables
file_path = os.path.dirname(open_file)
file_name = os.path.basename(open_file)


# Define unlock function
def unlock():
    # Check of the folder "tempfolder" in the selected file's path.
    # If folder does not exist, create it.
    if not os.path.exists(file_path + '/tempfolder/'):
        os.makedirs(file_path + '/tempfolder/')

    # Copy the file to unlock into tempfolder directory and unzip it.
    copyfile(file_path + '/' + file_name, file_path + '/tempfolder/' + file_name)

    with zipfile.ZipFile(file_path + '/tempfolder/' + file_name, 'r') as zip_ref:
        zip_ref.extractall(file_path + '/tempfolder/')

    # Loop through the files in tempfolder/xl/ directory after a file named 'workbook.xml'
    for file in os.listdir(file_path + '/tempfolder/xl/'):
        if file == 'workbook.xml':
            # Open and read the workbook.xml file. Encoding set to 'utf-8' to avoid problems with other languages.
            with open(file_path + '/tempfolder/xl/' + file, 'r', 'utf-8') as wb:
                Read_Workbook = wb.read()
                # Look for a the tag containing workbookProtection
                rm_wbprot_string = search('(?=<workbookProtection).*?(/>)', Read_Workbook)
                wb.close()
                # Check if the workbookProtection tag is found(i.e. the workbook is locked)
                if rm_wbprot_string:
                    # Replace the string with a 'nothing' string.
                    Remove_String = rm_wbprot_string.group(0)
                    wb_prot_removed = Read_Workbook.replace(Remove_String, "")

                    # Overwrite the locked workbook.xml with an unlocked version.
                    new_wb = open(wb.name, 'w+', 'utf-8')
                    new_wb.write(wb_prot_removed)
                    new_wb.close()

    # Loop through all the sheets contained in the workbook.
    for file in os.listdir(file_path + '/tempfolder/xl/worksheets/'):
        if file.lower().endswith('.xml'):
            # Open and read each sheet file. Encoding set to 'utf-8' to avoid problems with other languages.
            with open(file_path + '/tempfolder/xl/worksheets/' + file, 'r', 'utf-8') as sh:
                Read_Sheet = sh.read()
                # Look for the a tag containing sheetProtection
                rm_sheetprot_string = search('(?=<sheetProtection).*?(/>)', Read_Sheet)
                sh.close()
                # Check if the sheetProtection tag is found(i.e. the worksheet is locked)
                if rm_sheetprot_string:
                    # Replace the string with a 'nothing' string.
                    Remove_String = rm_sheetprot_string.group(0)
                    sheet_prot_removed = Read_Sheet.replace(Remove_String, "")

                    # Overwrite the locked sheet#.xml with an unlocked version.
                    new_sheet = open(sh.name, 'w+', 'utf-8')
                    new_sheet.write(sheet_prot_removed)
                    new_sheet.close()

    # Remove previously copied file from the tempfolder directory since
    # it will be in the way when putting the file back together.
    os.remove(file_path + '/tempfolder/' + file_name)
    # Split the name into File name and File Extension
    splitname = file_name.split(".")
    # Putting the file back together by creating a .ZIP archive
    make_archive(file_path + '/' + splitname[0], "zip", file_path + '/tempfolder/')
    # Renaming the .ZIP archive and changing it back to the original format.
    os.rename(file_path + '/' + splitname[0] + ".zip", file_path + '/' + splitname[0] + "_unlocked." + splitname[1])
    # Removing the tempfolder.
    rmtree(file_path + '/tempfolder/')


if __name__ == "__main__":
    try:
        unlock()
    except:
        print("Exit")
