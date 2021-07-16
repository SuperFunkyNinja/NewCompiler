import os
import fitz
import textwrap
import openpyxl
import sys
from pathlib import Path
from glob import glob
from datetime import datetime

# set index and log file names
EXCEL = Path("NewIndex.xlsx")
LOG = Path("LogFile.txt")
DESKTOP = Path("E:\Desktop")

# set working directory to python file location
WORKING = Path(__file__).parent.absolute()

# set up log file for output and error reporting with timestamp
now = datetime.now()
timestamp = str(now.strftime("%H:%M:%S on %d/%m/%Y"))
logFile = open(WORKING / LOG, "w")
logFile.write(f"This log file was created at {timestamp}.\n\n")

# check for excel file and throw error if not found
try:
    wb = openpyxl.load_workbook(WORKING / EXCEL, data_only=True)
except:
    logFile.write("Cannot find excel file.")
    sys.exit("Cannot find excel file.")

sheet = wb["Index"]  # create sheet object

# # check for blank pdf and throw error if not found
# PDF = Path(sheet["B3"].value)
# try:
#     newDoc = fitz.open(WORKING / PDF)
# except:
#     logFile.write("Cannot find pdf file.")
#     sys.exit("Cannot find pdf file.")

# pull global values from excel file
PROJECT = Path(sheet["B1"].value)
# OFFSET = sheet["B4"].value
# HEIGHT = sheet["B5"].value
# SIZE = sheet["B6"].value

# set output paths
OUTPUT = Path(str(sheet["B2"].value) + " OUTPUT.pdf")
OUTPUTERROR = Path("OUTPUT - ERROR check log file.pdf")

# remove leftover output files from previous iterations
try:
    os.remove(WORKING / LOG)
except:
    pass
try:
    os.remove(WORKING / OUTPUT)
except:
    pass
try:
    os.remove(WORKING / OUTPUTERROR)
except:
    pass

files = []  # empty list to store file references
pattern = "*.pdf"  # pattern for searching all the PDF in working directory

# search working directory and build list of all PDFs
for dir, _, _ in os.walk(PROJECT):
    files.extend(glob(os.path.join(dir, pattern)))

ref = 1  # index for iterating through doc refs
refs = {}  # dictionary for storing values from index file

# iterate through index rows and assign values to dictionary
for row in range(9, sheet.max_row + 1):
    front = sheet["D" + str(row)].value
    comment = sheet["E" + str(row)].value
    main = sheet["F" + str(row)].value

    if front is None:  # break out of loop if all rows read
        break

    refs.setdefault(
        ref,
        {"front": front, "comment": comment, "main": main},
    )
    ref = ref + 1

# create lists for file refs to be used and catching errors
numbers = []
for i in range(1, len(refs) + 1):
    numbers.append(refs[i]["front"])
    numbers.append(refs[i]["comment"])
    numbers.append(refs[i]["main"])

# create list of file references to check for missing or duplciate references
duplicates = []
missing = []

# check for unique references to be used going forward, and store missing refs and refs that can refer to multiple files
for i in numbers:
    count = 0
    for j in files:
        if i.lower() in j.lower():
            count += 1
            if count > 1:
                duplicates.append(i)
    if count == 0:
        missing.append(i)

# remove duplicate references from duplicates list
duplicates = list(dict.fromkeys(duplicates))

# write duplicates to log file
if len(duplicates) != 0:
    logFile.write("ERROR - The following duplicates were found:\n\n")
    for i in duplicates:
        logFile.write(i + "\n")
        for j in files:
            if i in j:
                logFile.write(j + "\n")
    logFile.write("\n")
    sys.exit("ERROR - Duplicates found.")

# write missing refs to log file
if len(missing) != 0:
    logFile.write("ERROR - The following references were not found:\n\n")
    for i in missing:
        logFile.write(i + "\n")
    logFile.write("\n")
    sys.exit("ERROR - Missing references.")

# # check that table of contents levels can be written
# for i in range(1, (len(refs)) + 1):
#     if i == 1 and refs[i]["lev"] != 1:
#         # check that TOC levels start at 1
#         logFile.write("ERROR - Table of contents level does not start at 1.")
#         sys.exit("ERROR - Table of contents level does not start at 1.")

#     # check that previous TOC level was higher or only jumped down 1
#     if i >= 2:
#         if (refs[i - 1]["lev"]) >= (refs[i]["lev"]) or (refs[i - 1]["lev"]) == (
#             refs[i]["lev"]
#         ) - 1:
#             pass
#         else:
#             logFile.write("ERROR - Table of contents level jumps down more than 1.")
#             sys.exit("ERROR - Table of contents level jumps down more than 1.")


logFile.write("This is the list of files (tab separated):\n\n")

print(files)

for i in refs:
    print(refs[i]["front"])
    print(refs[i]["comment"])
    print(refs[i]["main"])

    for j in files:
        if refs[i]["front"] in j:
            newDoc = fitz.open(WORKING / j)

    for j in files:
        if refs[i]["comment"] in j:
            newComment = fitz.open(WORKING / j)
            newDoc.insertPDF(newComment)

    for j in files:
        if refs[i]["main"] in j:
            newMain = fitz.open(WORKING / j)
            newDoc.insertPDF(newMain)

    newName = Path(refs[i]["main"] + ".pdf")

    logFile.write(refs[i]["main"] + "\n")

    newDoc.save(DESKTOP / newName)

# # create lists for PDF bookmarks, headings and page numbers
# # page numbers seperate from headings so they can be positioned seperately
# tocPDF = []
# tocSection = []
# tocRev = []
# tocPageNo = []
# tocHead = []
# tocRef = []

# fileerror = 0

# # create section heading page when needed
# def title_page(section, title):
#     blankPage = fitz.open(WORKING / PDF)  # create object from blank pdf
#     tempPage = blankPage[0]  # select page
#     p1 = fitz.Point(50, 400)  # set section heading position
#     t1 = f"Section {section}"  # section heading text
#     tempPage.insertText(p1, t1, fontsize=25)  # add section geading text
#     p2 = fitz.Point(50, 450)  # set section title position
#     wrapper = textwrap.TextWrapper(width=30)  # wrap section title text
#     t2 = wrapper.wrap(text=title)
#     tempPage.insertText(p2, t2, fontsize=25)  # add section title text
#     return blankPage


# for index in refs:
#     # build pdf bookmark table entry
#     entry = []
#     entry.append(refs[index]["lev"])
#     entry.append(f"Section {refs[index]['sect']} - {refs[index]['head']}")
#     entry.append(len(newDoc) + 1)
#     tocPDF.append(entry)  # add bookmarks list

#     # collect relevant info for the table of contents
#     contSection = f"Section {refs[index]['sect']}"
#     contRev = f"Revision: {refs[index]['rev']}"
#     contPageNo = f"Page {(len(newDoc) + 1 + OFFSET)}"
#     contHead = refs[index]["head"]
#     contRef = refs[index]["fil"]

#     # if table of contents entry required add it here
#     if refs[index]["tocs"] is not None and refs[index]["tocs"].lower() in "xyes":
#         tocSection.append(f"{contSection} - {contHead}")
#         tocPageNo.append(contPageNo)
#         # if revision number collumn required add it here
#         if refs[index]["rev"] is not None:
#             tocRev.append(contRev)
#             logFile.write(
#                 f"{contPageNo}\t{contSection}\t{contRef}\t{contRev}\t{contHead}\n"
#             )
#         else:  # if no revision entry then leave collumn blank
#             tocRev.append(" ")
#             logFile.write(f"{contPageNo}\t{contSection}\t{contRef}\t\t{contHead}\n")

#     # if section title page is required then add it here
#     if refs[index]["titl"] is not None and refs[index]["titl"].lower() in "xyes":
#         tpage = title_page(refs[index]["sect"], refs[index]["head"])
#         newDoc.insertPDF(tpage)

#     # if there is a file reference for this heading then add the file here
#     fil = refs[index]["fil"]
#     if fil is not None:
#         for i in files:
#             if fil.lower() in i.lower():
#                 try:  # try to pull the file from the harddrive
#                     newPages = fitz.open(i)
#                     newDoc.insertPDF(newPages)
#                     print("Inserted file " + fil)
#                 except:  # write reference to log if unsuccessful
#                     logFile.write(
#                         f"\n**** Cannot find a file on the local harddrive: {i} ****\n"
#                     )
#                     fileerror = 1
#                     print("Error finding file " + fil)

# try to add the bookmarks list to the pdf, add error message if unsuccessful
# tocerror = 0
# try:
#     newDoc.setToC(tocPDF)
# except:
#     logFile.write(
#         "\n**** Cannot write TOC to PDF, check you are not jumping down more than one level. ****\n"
#     )
#     tocerror = 1

# # turn the table of contents headings and page numbers lists into return separated strings
# tocSection = "\n".join(tocSection)
# tocRev = "\n".join(tocRev)
# tocPageNo = "\n".join(tocPageNo)

# # check there are no errors, write table of contents to first sheet
# if len(missing) == 0 and len(duplicates) == 0 and tocerror == 0 and fileerror == 0:
#     tocTitle = "Table of Contents"
#     p = fitz.Point(230, HEIGHT)  # This is the position of 'table of contents' (x, y)
#     p1 = fitz.Point(40, HEIGHT + 20)  # This is the position of the headings (x, y)
#     p2 = fitz.Point(450, HEIGHT + 20)  # This is the position of the revisions (x, y)
#     p3 = fitz.Point(500, HEIGHT + 20)  # This is the position of the page numbers (x, y)

#     newDoc[0].insertText(p, tocTitle, fontsize=15)
#     newDoc[0].insertText(p1, tocSection, fontsize=SIZE)
#     newDoc[0].insertText(p2, tocRev, fontsize=SIZE)
#     newDoc[0].insertText(p3, tocPageNo, fontsize=SIZE)

# if there were duplicate or missing refs, or table of contents errors, write warning message to first page
# else:
#     t = "**** Document not complete ****\n\n****** Please check log file ******"
#     p = fitz.Point(100, 300)
#     c = (1, 0, 0)
#     newDoc[0].insertText(p, t, fontsize=30, color=c)
#     OUTPUT = OUTPUTERROR

# # save the final pdf and close down working files
# try:
#     newDoc.save(WORKING / OUTPUT, garbage=4, deflate=1)
#     newDoc.close()
# except:
#     logFile.write("\n**** PDF is locked for editing! Cannot Create new PDF ****\n")
wb.close()
logFile.close()