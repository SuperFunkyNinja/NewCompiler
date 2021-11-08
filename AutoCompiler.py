import os
import fitz
import textwrap
import openpyxl
import sys
from pathlib import Path
from glob import glob
from datetime import datetime

# set index and log file names
LOG = Path("LogFile.txt")
DESKTOP = Path(r"C:\Users\design01\Desktop")
SOURCE = Path(r"C:\Users\design01\Desktop\Transmittal")

WORKING = DESKTOP

try:
    os.remove(WORKING / LOG)
except:
    pass

# set working directory to python file location
# WORKING = Path(__file__).parent.absolute()

# set up log file for output and error reporting with timestamp
now = datetime.now()
timestamp = str(now.strftime("%H:%M:%S on %d/%m/%Y"))
logFile = open(WORKING / LOG, "w")
logFile.write(f"This log file was created at {timestamp}.\n\n")

files = []  # empty list to store file references
pattern = "*.pdf"  # pattern for searching all the PDF in working directory

# search working directory and build list of all PDFs
for dir, _, _ in os.walk(SOURCE):
    files.extend(glob(os.path.join(dir, pattern)))


refs = []


for i in files:
    if "VDRL" in i:
        start = i.find("SLPRAM")
        end = i.find(" - ")
        refs.append(i[start:end])

print(refs)

for i in refs:
    print(i)
    found = ""
    for j in files:
        if i.lower() in j.lower() and "front" in j.lower():
            front = j
            found += " - front"
        if i.lower() in j.lower() and "crs" in j.lower():
            comment = j
            found += " - comment"
        if i.lower() in j.lower() and "vdrl" in j.lower():
            mainFile = j
            description = j[j.find(" - ") : j.find(" - ") + 20]
            found += " - main"

    if found != " - comment - front - main":
        logFile.write("missing reference: " + i + found + "/n")
        print("missing reference: " + i + found)

    newDoc = fitz.open(front)
    newComment = fitz.open(comment)
    newMain = fitz.open(mainFile)

    newDoc.insertPDF(newComment)
    newDoc.insertPDF(newMain)

    newName = Path("VDRL-" + i + description + ".pdf")

    logFile.write(str(newName) + "\n")

    newDoc.save(DESKTOP / newName)
