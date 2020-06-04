#!/usr/bin/env python3

# UPDATE THIS EVERY TIME A NEW RELEASE IS PACKAGED
VERSION = "1.5"

# Spyglass
# Source code by Derps aka Panzer Vier
# Modifications made by Khronion (KH)
# Ported to Python 3 with additional modifications by Zizou (Ziz)

import requests
import gzip
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.styles.colors import RED, GREEN, YELLOW
from bs4 import BeautifulSoup
import math
from datetime import datetime
from pathlib import Path
import sys

logpath = "debug.log"


# Method for writing a debug log
def write_log(text):
    with open(logpath, "a") as out:
        out.write(datetime.now().strftime(f"[%Y-%m-%d %H:%M:%S] {text}\n"))

# Method for getting user input
def query(text, options):
    while True:
        response = input(text)
        if response in options:
            return response

# Method for downloading data dump, and saving it to disk
def download_dump():
    dump_request = requests.get("https://www.nationstates.net/pages/regions.xml.gz", stream = True)
    with open("regions.xml.gz", "wb") as data_dump:
        for chunk in dump_request.iter_content(chunk_size = 16*1024):
            data_dump.write(chunk)

# Parse arguments, if any...

# Show help message and terminate
if "-h" in sys.argv or "--help" in sys.argv:
    print(f"Spyglass {VERSION}: Generate NationStates region update timesheets.\n")
    print("Developed by Panzer Vier, with additions by Khronion and Zizou\n")
    print(f"usage: {sys.argv[0]} [-h] [-n NATION] [-o OUTFILE] [-s | -l PATH]\n")
    print("""Optional arguments:
     -h           Show this help message and exit.
     -n NATION    Specify Nation to identify user by. In order to comply with
                  NationStates API rules, this must be the user's nation. Use
                  underscores instead of spaces.
     -o OUTFILE   File to output the generated timesheet in XLSX format to.
     -s           Suppress creating a debug log file. Log files are written to
                  the current working directory.
     -l PATH      Write debug log to specified path.
     -m           Generate a minimized sheet without WFEs and embassies
    """)
    print("""If run without arguments, Spyglass runs in interactive mode and outputs to its
working directory.""")
    sys.exit()

process_embassies = True
log = True

SpeedOverride = False
MinorTime = 5400
MajorTime = 9000

now = datetime.now()
YMD = f"{now.year}-{now.month}-{now.day}"

# Set nation name
if "-n" in sys.argv:
    UAgent = sys.argv[sys.argv.index("-n") + 1]
else:
    print(f"Spyglass {VERSION}: Generate NationStates region update timesheets.")
    UAgent = input("Nation Name: ")
    filename = f"SpyglassSheet{YMD}.xlsx"

    if query("Include region embassies? (y/n, defaults to y) ", ["y", "n", ""]) == "n":
        process_embassies = False

    # Ziz: Update lengths are now 1.5hrs and 2.5hrs for minor and major respectively
    if query("Do you want to manually specify update lengths? (y/n, defaults to n) ", ["y", "n", ""]) == "y":
        try:
            MinorTime = input("Minor Time, seconds (5400): ")
        except SyntaxError:
            MinorTime = 5400
        try:
            MajorTime = input("Major Time, seconds (9000): ")
        except SyntaxError:
            MajorTime = 9000
        SpeedOverride = True

# Set output filename
if "-o" in sys.argv:
    filename = sys.argv[sys.argv.index("-o") + 1]
else:
    filename = f"SpyglassSheet{YMD}.xlsx"

# Enable debug log
if "-s" in sys.argv:
    log = False

if "-m" in sys.argv:
    process_embassies = False
else:
    if "-l" in sys.argv:
        logpath = sys.argv[sys.argv.index("-l") + 1]
    starting_args = " ".join(sys.argv[1:])
    write_log(f"INFO Spyglass started with arguments: {starting_args}")
    write_log(f"INFO User Nation: {UAgent}")
    write_log(f"INFO Out File: {filename}")

# Set headers as required by NS TOS
headers = {"User-Agent": f"Spyglass-Fork/{VERSION} (developer:aptenodytezizou@gmail.com; user:{UAgent}; Authenticating)"}

# Verify specified nation is valid -- terminate if not
try:
    params = {"nation": UAgent, "q": "influence"}
    testreq = requests.get("https://www.nationstates.net/cgi-bin/api.cgi", headers = headers, params = params)
    testreq.raise_for_status()
    headers = {"User-Agent": f"Spyglass-Fork/{VERSION} (developer:aptenodytezizou@gmail.com; user:{UAgent})"}
except requests.exceptions.HTTPError:
    print("Nation not found. Be sure to input the name of a nation that actually exists.")
    if log:
        write_log(f"ERR  {UAgent} is not a valid nation. Terminating.")
    sys.exit()

if log:
    write_log(f"INFO Minor length: {MinorTime}")
    write_log(f"INFO Major length: {MajorTime}")

# Pulling a list of regions that are founderless and non-passworded. Eventually, we'll go through and highlight those
# on the sheet

if log:
    write_log("INFO Searching for data dump...")

# Total number of queries is low, so rate limit is unescessary
req = requests.get("https://www.nationstates.net/cgi-bin/api.cgi?q=regionsbytag;tags=-password", headers = headers)
req2 = requests.get("https://www.nationstates.net/cgi-bin/api.cgi?q=regionsbytag;tags=founderless", headers = headers)
html = req.text
html2 = req2.text

# Ziz: If a data dump is detected in the current directory, ask if user wants to re-download latest dump
# Ziz: Otherwise just download the latest data dump if nothing is detected
dump_path = Path("./regions.xml.gz")
if dump_path.exists() and dump_path.is_file():
    if query("Existing data dump found. Do you want to re-download the latest dump? (y/n, defaults to y) ", ["y", "n", ""]) == "y":
        if log:
            write_log("INFO Found data dump, but re-downloading the latest..")
        print("Pulling data dump...")
        download_dump()
        if log:
            write_log("INFO Download complete!")
    else:
        write_log("INFO Using data dump already present...")
        print("Using current dump...")
else:
    if log:
        write_log("INFO No existing data dump found, downloading latest...")
    print("No existing data dump found. Pulling data dump...")
    download_dump()
    if log:
        write_log("INFO Download complete!")

redFill = PatternFill(start_color = RED, end_color = RED, fill_type = "solid")
greenFill = PatternFill(start_color = GREEN, end_color = GREEN, fill_type = "solid")
yellowFill = PatternFill(start_color = YELLOW, end_color = YELLOW, fill_type = "solid")

# Un-gzipping
# Ziz: Now we can just decompress the dump and hand it to the parser without writing to disk
with gzip.open("regions.xml.gz", "rb") as dump:
    regions = dump.read()

# Opening up my virtual sheet. Gotta find a better name for it, sometime. The pink tab colour's pretty sweet, tho.
wb = Workbook()
ws = wb.active
ws.title = "Spyglass Timesheet"
ws.sheet_properties.tabColor = "FFB1B1"

RegionList = list()
RegionURLList = list()
RegionWFEList = list()
RegionEmbassyList = list()
NumNationList = list()
DelVoteList = list()
ExecList = list()
MajorList = list()

# Sanitize our founderless regions list a wee bit, 'cause at the moment, it's xml, and xml is gross.
print("Processing data...")
UnfoundedList = BeautifulSoup(html2, "lxml-xml").find("REGIONS").string.split(",")
PWlessList = BeautifulSoup(html, "lxml-xml").find("REGIONS").string.split(",")

# Pulling, in order, region names, converting to a region url, number of nations in that region, and voting power that
# delegate has.

# Ziz: Data dump soup, mmm... almost as tasty as people!
data_dump_soup = BeautifulSoup(regions, "lxml-xml")
for name in data_dump_soup.find_all("NAME"):
    RegionList.append(name.string)
    UrlString = f"=HYPERLINK(\"https://www.nationstates.net/region={name.string}\")"
    RegionURLList.append(UrlString.replace(" ", "_"))
for nation_amount in data_dump_soup.find_all("NUMNATIONS"):
    NumNationList.append(int(nation_amount.string))
for del_votes in data_dump_soup.find_all("DELEGATEVOTES"):
    DelVoteList.append(int(del_votes.string))
for auth in data_dump_soup.find_all("DELEGATEAUTH"):
    AuthString = auth.string
    if AuthString[0] == "X":
        ExecList.append(True)
    else:
        ExecList.append(False)

# KH: pull major times from daily dump
for major_time in data_dump_soup.find_all("LASTUPDATE"):
    MajorList.append(int(major_time.string))

# KH: gather WFE info
for wfe in data_dump_soup.find_all("FACTBOOK"):
    text = wfe.string
    try:
        if text[0] in ["=", "+", "-", "@"]:
            text = f"'{text}" # IMPORTANT: prevent excel from parsing WFEs as formulas
        RegionWFEList.append(text)
    except TypeError: # no WFE
        RegionWFEList.append("")

for region_embassies in data_dump_soup.find_all("EMBASSIES"):
    embassies = list()
    if process_embassies:
        for embassy in region_embassies.find_all("EMBASSY"):
            embassies.append(embassy.string)
    RegionEmbassyList.append(",".join(embassies))


# Grabbing the cumulative number of nations that've updated by the time a region has.
# The first entry is zero because time calculations need to reflect the start of region update, not the end
CumulNationList = [0]
for a in NumNationList:
    CumulNationList.append(int(a) + CumulNationList[-1])

# Calculate speed based on total population
CumulNations = CumulNationList[-1]
MinorNatTime = float(MinorTime) / CumulNations
MajorNatTime = float(MajorTime) / CumulNations
MinTime = list()
MajTime = list()

# Getting the approximate major/minor update times.
for a in CumulNationList:
    temptime = int(a * MinorNatTime)
    tempsecs = temptime % 60
    tempmins = int(math.floor(temptime / 60) % 60)
    temphours = int(math.floor(temptime / 3600))
    MinTime.append(f"{temphours}:{tempmins}:{tempsecs}")

# If user specifies update length, use special handling.
if SpeedOverride:
    for a in CumulNationList:
        temptime = int(a * MajorNatTime)
        tempsecs = temptime % 60
        tempmins = int(math.floor(temptime / 60) % 60)
        temphours = int(math.floor(temptime / 3600))
        MajTime.append(f"{temphours}:{tempmins}:{tempsecs}")
else:
    for a in MajorList:
        temptime = a - MajorList[0]
        tempsecs = temptime % 60
        tempmins = int(math.floor(temptime / 60) % 60)
        temphours = int(math.floor(temptime / 3600))
        MajTime.append(f"{temphours}:{tempmins}:{tempsecs}")

# Splashing some headers and stuff onto the spreadsheet for legibility purposes!
ws["A1"].value = "Regions"
ws["B1"].value = "Region Link"
ws["C1"].value = "# Nations"
ws["D1"].value = "Tot. Nations"
ws["E1"].value = "Minor Upd. (est)"
ws["F1"].value = "Major Upd. (true)"
ws["G1"].value = "Del. Votes"
ws["H1"].value = "Del. Endos"
if process_embassies:
    ws["I1"].value = "Embassies"
ws["J1"].value = "WFE"

ws["L1"].value = "World "
ws["M1"].value = "Data"
ws["L2"].value = "Nations"
ws["L3"].value = "Last Major"
ws["L4"].value = "Secs/Nation"
ws["L5"].value = "Nations/Sec"
ws["L6"].value = "Last Minor"
ws["L7"].value = "Secs/Nation"
ws["L8"].value = "Nations/Sec"
ws["L10"].value = "Spyglass Version"
ws["L11"].value = "Date Generated"
ws["M2"].value = CumulNations
ws["M3"].value = MajorList[-1] - MajorList[0]
ws["M4"].value = float(MajorList[-1] - MajorList[0]) / CumulNations
ws["M5"].value = 1 / (float(MajorList[-1] - MajorList[0]) / CumulNations)
ws["M6"].value = MinorTime
ws["M7"].value = MinorNatTime
ws["M8"].value = 1 / MinorNatTime
ws["M10"].value = VERSION
ws["M11"].value = YMD

# There's probably a better way of doing this, but my coding skills are dubious :^)
# Anyways, actually pasting the information from our various lists into the spreadsheet.
counter = 0

for a in RegionList:
    # If a region's founderless, highlight it for easy reference. Add a tilde, 'cause my spreadsheet program doesn't
    # do filtering by colour

    b = a
    # KH: ~ indicates hittable
    # KH: yellow = passwordless and exec delegate
    if a in PWlessList and ExecList[counter]:
        ws.cell(row = counter + 2, column = 1).fill = yellowFill
        ws.cell(row = counter + 2, column = 2).fill = yellowFill
        b = f"{a}~"
    # KH: green = founderless and passwordless
    if a in UnfoundedList and a in PWlessList:
        ws.cell(row = counter + 2, column = 1).fill = greenFill
        ws.cell(row = counter + 2, column = 2).fill = greenFill
        b = f"{a}~"
    # KH: red = passwordless
    if a not in PWlessList:
        ws.cell(row = counter + 2, column = 1).fill = redFill
        ws.cell(row = counter + 2, column = 2).fill = redFill
        b = f"{a}*"
    ws.cell(row = counter + 2, column = 1).value = b
    ws.cell(row = counter + 2, column = 2).value = RegionURLList[counter]
    ws.cell(row = counter + 2, column = 3).value = NumNationList[counter]
    ws.cell(row = counter + 2, column = 4).value = CumulNationList[counter]
    ws.cell(row = counter + 2, column = 5).value = MinTime[counter]
    ws.cell(row = counter + 2, column = 5).alignment = Alignment(horizontal = "right")
    ws.cell(row = counter + 2, column = 6).value = MajTime[counter]
    ws.cell(row = counter + 2, column = 6).alignment = Alignment(horizontal = "right")
    ws.cell(row = counter + 2, column = 7).value = DelVoteList[counter]
    ws.cell(row = counter + 2, column = 8).value = DelVoteList[counter] - 1
    ws.cell(row = counter + 2, column = 9).value = RegionEmbassyList[counter]
    ws.cell(row = counter + 2, column = 10).value = RegionWFEList[counter]
    ws.cell(row = counter + 2, column = 11).value = " "

    # Highlight delegate-less regions. They're good for tagging, or whatever~
    if DelVoteList[counter] == 0:
        ws.cell(row=counter + 2, column=8).fill = redFill
    counter += 1

# You know those situations where you can't quite get code to work, and kinda fumble around until you find something
# that does?
#
# I'm 90% sure this isn't the way to do it, but I couldn't get it working otherwise.
# Anyways, setting the region name column's width, so that it doesn't cut everything off.
sheet = wb["Spyglass Timesheet"]
sheet.column_dimensions["A"].width = 45
sheet["J1"].alignment = Alignment(horizontal = "right")

if log:
    write_log("INFO Done processing data! Saving sheet.")

# Really should just name the sheets 'Derps is amazing in every conceivable way'. Would be some free ego-massage.
print("Saving sheet...")
wb.save(filename)
# Ziz: Don't delete the data dump, since it can be reused if it's recent enough

if log:
    write_log(f"INFO Successfully saved to {filename}")
    write_log(f"INFO Spyglass run complete. Exiting...")

sys.exit()
