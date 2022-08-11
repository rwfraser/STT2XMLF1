"""
Current version as of July 8/22 => 5.4, scrapes the most recently used LOCATION from myearringdepot.com
Current version as of March 3/22 => 5.3, has PASSED preprocessing and PRODUCTION.  See STT2XMLDoc.py for project documentation
Current version as of March 2/22 => 5.2 is production, 5.3 is in development, has PASSED preprocessing.  See STT2XMLDoc.py for project documentation

NOTES:  STT.txt file must be utf-8 encoded, single spaced. 
"""

from sys import argv
from os.path import exists
from xml.dom.minidom import parseString
import MEDGlobals
import os
import pathlib
import shutil
import pysftp
import paramiko
import time
from LocationAdder import NextLocation
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import requests
import json
imprt base64

# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
import requests
from requests_html import HTMLSession
from urllib.parse import urljoin
from webdriver_manager.chrome import ChromeDriverManager


# ********** Run time configuration: ************************
# LOCATION = 'Bc3o05'
PRE_PROCESSOR = False
DEBUG = False
DEBUG_WAIT = 7
RUN_FUTURE_CODE = False
LIVE = True
LIVE_WAIT = 5

base_url = "https://rogeridaho.wpengine.com/wp-content/uploads/"
base_path = "C:\\Users\\RogerIdaho"

XML_HEADER = '<?xml version="1.0" encoding="UTF-8" ?>'
CURRENT_VERSION = argv[0][:-3]
XML_COPYRIGHT_VERSION_NOTICE = (
    "<!-- "
    + CURRENT_VERSION
    + "Copyright 2022 myearringdepot.com All Rights Reserved  -->"
)
OPENING_PRODUCT_TAG = "<Product>"
SKU_String = ""
LOCATION_String = ""
ITEM_OUTPUT_STRING = ""
OUTPUT_STRING = ""
SixBitOpeningTags = "<Items>"
SixBitClosingTags = "</Items>"
SixBitOutputString = ""
ConcatenatedTags = ""
ConcatenatedMaterials = ""
SixBitOccasion = ""
EtsyOccasion = ""
CumulativeTermCount = 0
CurrentItemTermCount = 0
ItemCounter = 1

YEAR = "22" #correct year should be derived from date/time stamp of DCIM folder on SD card.  
YR = "22"
YEAR_PREFIX = "20"
FULL_YEAR = YEAR_PREFIX + YEAR
LOCAL_PATH_ROOT = "C:\\users\\RogerIdaho\\"
SFTP_SERVER = "rogeridaho.sftp.wpengine.com"
SFTP_ROOT = "./wp-content/uploads"
SFTP_ACCOUNT = "rogeridaho"
SFTP_PASSWORD = "Cyb4rS4cur1ty!!##"
SFTP_PORT = 2222
SFTP_PROTOCOL = "SFTP"
PHOTOS_PER_ITEM = 5
ITEM_PER_TRAY = 15


Shop_Section_Names_List = [
    "Chandelier Earrings",
    "Clip On Earrings",
    "Cuff and Wrap Earrings",
    "Dangle and Drop Earrings",
    "Ear Jackets",
    "Hoop Earrings",
    "Screwback Earrings",
    "Stud Earrings",
    "Threader Earrings",
    "French Hooks",
    "Single Earrings",
    "Stud Dangles",
]

Shop_Section_IDs_List = []

TagsParametersList = [
    "Theme1",
    "Theme2",
    "Style",
    "era",
    "Color1",
    "Color2",
    "Gemstone1",
    "Gemstone2",
    "Metal1",
    "Metal2",
    "Mineral1",
    "Mineral2",
    "Material1",
    "Material2",
    "Wearer1",
    "Wearer2",
    "Occasion1",
    "Occasion2",
    "Character1",
    "Character2",
    "Brand1",
    "Brand2",
    "Tag1",
    "Tag2",
]

MaterialsParametersList = [
    "Gemstone1",
    "Gemstone2",
    "Metal1",
    "Metal2",
    "Mineral1",
    "Mineral2",
    "Material1",
    "Material2",
]

EtsyDescriptionParametersList = [
    "Type",
    "Style",
    "Theme1",
    "Theme2",
    "Style",
    "era",
    "Color1",
    "Color2",
    "Gemstone1",
    "Gemstone2",
    "Metal1",
    "Metal2",
    "Mineral1",
    "Mineral2",
    "Material1",
    "Material2",
    "Wearer1",
    "Wearer2",
    "Occasion1",
    "Occasion2",
    "Character1",
    "Character2",
    "Brand1",
    "Brand2",
    "Tag1",
    "Tag2",
    "Gemstone1",
    "Gemstone2",
    "Metal1",
    "Metal2",
    "Mineral1",
    "Mineral2",
    "Material1",
    "Material2",
]

Earring_Parameters_Dict = {
    "PR": "Price",
    "PT": "PriceTier",
    "QT": "Quantity",
    "TT": "Title",
    "DL": "Description",
    "DS": "ShortDescription",
    "TY": "Type",
    "C1": "Color1",
    "see one": "Color1",
    "sea one": "Color1",
    "C2": "Color2",
    "C to": "Color2",
    "C too": "Color2",
    "C two": "Color2",
    "sea two": "Color2",
    "see two": "Color2",
    "ME1": "Metal1",
    "ME 1": "Metal1",
    "ME one": "Metal1",
    "Me1": "Metal1",
    "ME2": "Metal2",
    "ME 2": "Metal2",
    "ME to": "Metal2",
    "ME too": "Metal2",
    "ME two": "Metal2",
    "Me2": "Metal2",
    "GE1": "Gemstone1",
    "GE 1": "Gemstone1",
    "GE one": "Gemstone1",
    "Ge1": "Gemstone1",
    "GE2": "Gemstone2",
    "GE 2": "Gemstone2",
    "GE to": "Gemstone2",
    "GE too": "Gemstone2",
    "GE two": "Gemstone2",
    "Ge2": "Gemstone2",
    "MA1": "Material1",
    "MA 1": "Material1",
    "MA one": "Material1",
    "M A1": "Material1",
    "Ma1": "Material1",
    "MA2": "Material2",
    "MA 2": "Material2",
    "MA to": "Material2",
    "MA too": "Material2",
    "MA two": "Material2",
    "Ma2": "Material2",
    "MI1": "Mineral1",
    "MI 1": "Mineral1",
    "MI one": "Mineral1",
    "Mi1": "Mineral1",
    "MI2": "Mineral2",
    "MI 2": "Mineral2",
    "MI to": "Mineral2",
    "MI too": "Mineral2",
    "MI two": "Mineral2",
    "Mi2": "Mineral2",
    "DH": "Height",
    "DW": "Width",
    "DD": "Diameter",
    "ER": "Era",
    "ERA": "Era",
    "era": "Era",
    "CN": "Condition",
    "ST": "Style",
    "WE1": "Wearer1",
    "WE 1": "Wearer1",
    "WE one": "Wearer1",
    "We1": "Wearer1",
    "WE2": "Wearer2",
    "WE 2": "Wearer2",
    "WE to": "Wearer2",
    "WE too": "Wearer2",
    "WE two": "Wearer2",
    "We2": "Wearer2",
    "TH1": "Theme1",
    "TH 1": "Theme1",
    "TH one": "Theme1",
    "Th1": "Theme11",
    "TH2": "Theme2",
    "TH 2": "Theme2",
    "TH to": "Theme2",
    "TH too": "Theme2",
    "TH two": "Theme2",
    "Th2": "Theme2",
    "OC1": "Occasion1",
    "OC 1": "Occasion1",
    "OC one": "Occasion1",
    "Oc1": "Occasion1",
    "OC2": "Occasion2",
    "OC 2": "Occasion2",
    "OC to": "Occasion2",
    "OC too": "Occasion2",
    "OC two": "Occasion2",
    "Oc2": "Occasion2",
    "CH1": "Character1",
    "CH 1": "Character1",
    "CH one": "Character1",
    "Ch1": "Character1",
    "CH2": "Character2",
    "CH 2": "Character2",
    "CH to": "Character2",
    "CH too": "Character2",
    "CH two": "Character2",
    "Ch2": "Character2",
    "BR1": "Brand1",
    "BR 1": "Brand1",
    "BR one": "Brand1",
    "Br1": "Brand1",
    "BR2": "Brand2",
    "BR 2": "Brand2",
    "BR to": "Brand2",
    "BR too": "Brand2",
    "BR two": "Brand2",
    "Br2": "Brand2",
    "SH1": "Shape1",
    "SH 1": "Shape1",
    "SH one": "Shape1",
    "Sh1": "Shape1",
    "SH2": "Shape2",
    "SH 2": "Shape2",
    "SH to": "Shape2",
    "SH too": "Shape2",
    "SH two": "Shape2",
    "Sh2": "Shape2",
    "TG1": "Tag1",
    "TG 1": "Tag1",
    "TG one": "Tag1",
    "TG one": "Tag1",
    "Tg1": "Tag1",
    "TG2": "Tag2",
    "TG 2": "Tag2",
    "TG to": "Tag2",
    "TG too": "Tag2",
    "TG two": "Tag2",
    "Tg2": "Tag2",
}

Current_Item_Parameters_Dict = {
    "Price": "",
    "PriceTier": "",
    "Quantity": "",
    "Title": "",
    "Description": "",
    "ShortDescription": "",
    "Type": "",
    "Color1": "",
    "Color2": "",
    "Metal1": "",
    "Metal2": "",
    "Gemstone1": "",
    "Gemstone2": "",
    "Material1": "",
    "Material2": "",
    "Mineral1": "",
    "Mineral2": "",
    "Height": "",
    "Width": "",
    "Diameter": "",
    "Era": "",
    "era": "",
    "Condition": "",
    "Style": "",
    "Wearer1": "",
    "Wearer2": "",
    "Theme1": "",
    "Theme2": "",
    "Occasion1": "",
    "Occasion2": "",
    "Character1": "",
    "Character2": "",
    "Brand1": "",
    "Brand2": "",
    "Shape1": "",
    "Shape2": "",
    "Tag1": "",
    "Tag2": "",
}

Section_Names_By_Type_Dict = {
    "French Hook": "Chandelier Earrings",
    "French hook": "Chandelier Earrings",
    "Clip": "Clip On Earrings",
    "Clips": "Clip On Earrings",
    "Cuff and Wrap Earrings": "Cuff and Wrap Earrings",
    "Ear Wire": "French Hook",
    "Post Dangle": "Dangle and Drop Earrings",
    "Ear Jackets": "Ear Jackets",
    "Ear Cuff": "Cuff and Wrap Earrings",
    "Hoop": "Hoop Earrings",
    "Screwback": "Screwback Earrings",
    "Screw Back": "Screwback Earrings",
    "Stud": "Stud Earrings",
    "Threader Earrings": "Threader Earrings",
    "French Hook": "French Hooks",
    "Lever Back": "Lever Back",
    "lever back": "Lever Back",
    "Leverback": "Lever Back",
    "Single": "Single Earrings",
    "Post": "Stud Dangles",
    "Post No Dangle": "Dangle and Drop Earrings",
}

Type_By_Section_Names_Dict = {
    "Chandelier Earrings": "French Hook",
    "Clip On Earrings": "Clip",
    "Cuff and Wrap Earrings": "Cuff and Wrap Earrings",
    "Dangle and Drop Earrings": "Post Dangle",
    "Ear Jackets": "Ear Jackets",
    "Hoop Earrings": "Hoop",
    "Screwback Earrings": "Screwback",
    "Stud Earrings": "Stud",
    "Threader Earrings": "Threader Earrings",
    "French Hooks": "French Hook",
    "Single Earrings": "Single",
    "Stud Dangles": "Post",
}

Section_IDs_By_Type_Dict = {
    "Post Dangle": "1208",
    "Post No Dangle": "1203",
    "Clip": "1205",
    "Clips": "1205",
    "Cluster": "1206",
    "French Hook": "1208",
    "Chandelier": "1204",
    "Hoop": "1212",
    "Screwback": "1213",
    "Leverback": "1213",
    "Adjustable Leverback": "1213",
    "Screw Back": "1213",
    "Lever Back": "1213",
    "Adjustable Lever Back": "1213",
    "Stud": "1214",
    "Post": "1208",
    "Earwire": "1203",
    "Ear Wire": "1203",
}

Etsy_Type_ID_Dict = {
    "Earrings": "1203",
    "Chandelier": "1204",
    "Clip On": "1205",
    "Cluster": "1206",
    "Cuff & Wrap": "1207",
    "Dangle & Drop": "1208",
    "Ear Jackets & Climbers": "2900",
    "Ear Weights": "1210",
    "Gauge & Plug": "1211",
    "Hoop Earrings": "1212",
    "Jhumka": "12185",
    "Kaan Chains": "1223",
    "Screw Back Earrings": "1213",
    "Stud Earrings": "1214",
    "Threader Earrings": "1215",
    "Earwire": "1203",
}


# LOCATION = (input("LOCATION START: "))  #  use SSH WP DB QUERY to retrieve last LOCATION from myearringdepot.com
if PRE_PROCESSOR or LIVE:
    user_agent = {"User-agent": "Mozilla/5.0"}
    SortedMEDPage = requests.get(
        "https://myearringdepot.com/?swoof=1&orderby=date", headers=user_agent
    ).text
    SortedMedPageString = str(SortedMEDPage)
    """
    driver.close()
    driver.quit()
    """
    Start_index = SortedMEDPage.find("data-product_sku") + 24
    End_index = Start_index + 6
    MostRecentLOCATION = SortedMEDPage[Start_index:End_index]
    print(f"Most recently used LOCATION is: {MostRecentLOCATION}")
    LOCATION = MostRecentLOCATION

if PRE_PROCESSOR:
    print("Sample run ONLY, be sure SD card is present in USB port")

if PRE_PROCESSOR:
    print(f"Using {NextLocation(LOCATION)} as start LOCATION")

if LIVE:
    print(
        f"We're LIVE! Verify starting LOCATION is:  {NextLocation(LOCATION)}, CTRL-Z to exit "
    )
    time.sleep(DEBUG_WAIT)

if RUN_FUTURE_CODE:

    def getLastLocationFromDB(location):
        """returns string of Last Used Location using sql command to retrieve all LOCATION strings, most recent first:"""
        # Database details
        # check Wordpress wp-config.php file in root folder of website and WP Engine's  SSH Gateway documentation
        DB_NAME = "wp_rogeridaho"
        DB_USER = "rogeridaho"
        DB_PASSWORD = "IK4k8wYuYasK9slwIIT8"
        DB_HOST = "127.0.0.1:3306"  # URL: https://phpmyadmin.wpengine.com/signon.php, https://pod-140622.wpengine.com
        DB_HOST_SLAVE = ""
        DB_PORT = 3306
        DB_CHARSET = "utf-8"
        DB_COLLATE = "utf-8_unicode_ci"
        TABLE_PREFIX = "wp_"  # php statements: $table_prefix = 'wp_';
        # setup SSH Gateway/connection string
        SQL_COMMAND = "SELECT meta_value FROM `wp_postmeta` WHERE `meta_key` = 'location' ORDER BY `wp_postmeta`.`meta_value` DESC"  # returns a list, first element is Last used Location string
        NextLocationPerDB = SQL_COMMAND[0]
        return NextLocationPerDB


def ProgressDots(UploadCounter):
    """Returns string. Prints one additional . each time a file is uploaded"""
    DotString = "."
    for i in range(UploadCounter):
        DotString = DotString + "."
    return DotString


# ********************* READ SD CARD ***********************
# SD card defaults to F: drive unless operator overrides:
LogicalDrive = ""
try:
    LogicalDrive = sys.argv[1]
    LogicalDrive = LogicalDrive.strip()
    if len(LogicalDrive) != 2:
        LogicalDrive = LogicalDrive + ":"

except:
    LogicalDrive = "G:"

if PRE_PROCESSOR:
    print(f"Using drive {LogicalDrive}")

#  We have no error checking so there can be ONLY the DCIM folder with ONE photo subfolder AND only one batch per day TFN  Feb 21/22
if RUN_FUTURE_CODE:
    Filename = Function_to_error_check_SD_Card(LogicalDrive)
    """ returns Filename as string => conducts error checking and verification for previous batches on the current date  and to locate the most recent photo filder on the SD card """


# ************ GET Filename  *******************************
photodirectory = "DCIM"  # this should be '\DCIM' but will need testing to verify
ImgRoot = LogicalDrive + photodirectory
for DCIMFolderName in os.scandir(ImgRoot):
    DirContents = DCIMFolderName
DCIMFolderName = str(DirContents)
DCIMFolderName = DCIMFolderName[11:-2]
if PRE_PROCESSOR:
    print(f"Photo folder name is: {DCIMFolderName}")

#  GET DATE AND BATCH
# Filename also provides the date.  Use it to create destination folders and copy/upload the photo files
Month = DCIMFolderName[4:6]
Day = DCIMFolderName[6:8]
month_string = DCIMFolderName[4:6]
day_string = DCIMFolderName[6:8]
year_string = YEAR  #

# check C: drive for previous batches from the same day => future development, for now batch_number = '01'
batch_number = "1"
batch_number = "0" + batch_number
batch_date = FULL_YEAR + "_" + month_string + "_" + day_string
if PRE_PROCESSOR:
    print(f"Run Date is: {batch_date}")

# **************** DETERMINE DESTINATION PATHS ***********************************
# Derive paths for local and SFTP storage
ImgFolder = LogicalDrive + "\\" + photodirectory + "\\" + DCIMFolderName
CameraPrefix = "100"
LocalDestination = LOCAL_PATH_ROOT
SFTPDestination = (
    SFTP_ROOT
    + "/"
    + YEAR_PREFIX
    + YEAR
    + "/"
    + Month
    + "/"
    + Day
    + "/"
    + batch_number
    + "/"
)
LocalDestination = (
    LOCAL_PATH_ROOT
    + YEAR_PREFIX
    + YEAR
    + "\\"
    + CameraPrefix
    + "-"
    + Month
    + Day
    + "-"
    + batch_number
)

if PRE_PROCESSOR:
    print(f"SFTP Destination is:  {SFTPDestination}")
    print(f"Local Destination is: {LocalDestination}")

#  *************** VERIFY # OF PHOTOS
# Verify number of photos is correct
Count = 0
if PRE_PROCESSOR:
    print(f"Path to photos is: {ImgFolder}")
    time.sleep(DEBUG_WAIT)
for filename in os.scandir(ImgFolder):
    if filename.is_file():
        Count += 1
if PRE_PROCESSOR:
    print(f"Number of files is: {Count}")
    if Count % PHOTOS_PER_ITEM != 0:
        print("Wrong number of photos")
BatchQuantity = Count / PHOTOS_PER_ITEM
if PRE_PROCESSOR:
    print(f"# Earrings = {BatchQuantity}")

# loop through files
PhotoFileList = []
PhotoFileList = os.listdir(ImgFolder) # 07/08/22 thumbs.db is present in PhotoFileList, but not in DOS
os.chdir(ImgFolder)
if PRE_PROCESSOR:
    for file in PhotoFileList:
        print(file, end="")
    else:
        pass

# ****************** CREATE DESTINATION FOLDERS AND COPY FILES ****************************
if LIVE:
    os.mkdir(LocalDestination)
for file in PhotoFileList:
    localpath = ImgFolder + "\\" + file
    if PRE_PROCESSOR:
        print(f"Path to source images is: {ImgFolder}")

    CDrivePath = LocalDestination + "\\" + file
    if PRE_PROCESSOR:
        print(f"Source file is: {localpath}")
        print(f"Destination path is: {CDrivePath}")
    if LIVE:
        shutil.copy(
            localpath, LocalDestination
        )  # copy files from SD card to local drive

#  SFTP: connect to sftp, create destination folders, upload the files from the SD
if LIVE or PRE_PROCESSOR:
    host = SFTP_SERVER
    port = SFTP_PORT
    transport = paramiko.Transport((host, port))
    password = SFTP_PASSWORD
    username = SFTP_ACCOUNT
    transport.connect(username=username, password=password)
    if transport.connect and PRE_PROCESSOR:
        print("SFTP Connection established")

sftp = paramiko.SFTPClient.from_transport(transport)
DestinationFolder = "./wp-content/uploads/" + FULL_YEAR + "/" + Month
if PRE_PROCESSOR:
    print(f"Destination folder => year + month is:  {DestinationFolder}")
    time.sleep(DEBUG_WAIT)
if LIVE:
    try:
        sftp.mkdir(DestinationFolder)
        DestinationFolder = DestinationFolder + "/" + Day
        sftp.mkdir(DestinationFolder)
        DestinationFolder = DestinationFolder + "/" + batch_number
    except:
        DestinationFolder = DestinationFolder + "/" + Day
        sftp.mkdir(DestinationFolder)
        DestinationFolder = DestinationFolder + "/" + batch_number


if LIVE:
    sftp.mkdir(DestinationFolder)

UploadCounter = 0
# iterate through ImgFolder List to upload, concatenate paths for each file
for file in PhotoFileList:
    path = SFTPDestination + file
    localpath = ImgFolder + "\\" + file
    if PRE_PROCESSOR:
        print(f"SFTP Dest is: {path}, Source is: {localpath}")
    if LIVE:
        sftp.put(localpath, path)
        print(
            f"{ProgressDots(UploadCounter)}"
        )  # call progress function in f string of print command to print one extra . for each file
        UploadCounter += 1
sftp.close()
transport.close()

if PRE_PROCESSOR:
    print("PREPROCESSING: Upload  NOT done.")
elif LIVE:
    print("SFTP Upload Complete")
time.sleep(DEBUG_WAIT)

# generate image number strings for photo paths/URLs
if LIVE or PRE_PROCESSOR:
    batch_size = BatchQuantity  # batch_size is legacy variable
else:
    batch_size = 30
imagenumberstrings = []
batch_size_integer = int(batch_size)
for i in range(1, (batch_size_integer * 5) + 1):
    j = str(i)
    if i < 10:
        j = "000" + j
    elif i > 9 and i < 100:
        j = "00" + j
    else:
        j = "0" + j
    imagenumberstrings.append(j)

# *************** INGEST and PARSE batch_dateSTT.txt file
INPUT_FILE = "C:\\users\\rogeridaho\\" + batch_date + "STT.txt"
if PRE_PROCESSOR:
    print(f"STT.txt filename is: {INPUT_FILE}")

with open(INPUT_FILE) as f:
    # INPUT_PARAMETER_TEMP_LIST_TEMP = f.readlines()
    # INPUT_PARAMETER_TEMP_LIST = [x[:-1] for x in INPUT_PARAMETER_TEMP_LIST_TEMP]
    INPUT_PARAMETER_TEMP_LIST = f.readlines()

# ***********************MAIN LOOP *********************************************
for s in INPUT_PARAMETER_TEMP_LIST:
    tempterm = INPUT_PARAMETER_TEMP_LIST[CumulativeTermCount]
    tempterm = tempterm.strip()
    ShortKeyList = []
    if tempterm != "and;" and tempterm != "end;":
        colon_position = INPUT_PARAMETER_TEMP_LIST[CumulativeTermCount].find(":")
        ShortKey = tempterm[0:colon_position]
        ShortKeyList.append(ShortKey)
        Value1 = tempterm[colon_position + 1 : len(tempterm)]
        StructuredPair = ""
        for key in Earring_Parameters_Dict:
            if key == ShortKey:
                CurrentItemTermCount = CurrentItemTermCount + 1
                if Earring_Parameters_Dict[key] == "Title":
                    Value1 = Value1.title()
                    Current_key = Earring_Parameters_Dict[key]
                    StructuredPair = Earring_Parameters_Dict[key] + ":" + Value1
                if Earring_Parameters_Dict[key] == "Occasion1":
                    SixBitOccasion = Value1
                    Current_key = Earring_Parameters_Dict[key]
                    StructuredPair = Earring_Parameters_Dict[key] + ":" + Value1
                #                 if Earring_Parameters_Dict[key] == 'Description'  fork Value1 for differing descriptions on website and Etsy
                if Earring_Parameters_Dict[key] == "Height":
                    TempValue1 = ""
                    Value1 = Value1.strip()
                    Value1 = Value1[0:5]
                    for x in Value1:
                        if x.isdigit() or x == ".":
                            TempValue1 = TempValue1 + x
                    Value1 = TempValue1
                    Current_key = Earring_Parameters_Dict[key]
                if Earring_Parameters_Dict[key] == "Width":
                    TempValue1 = ""
                    Value1 = Value1.strip()
                    Value1 = Value1[0:5]
                    for x in Value1:
                        if x.isdigit() or x == ".":
                            TempValue1 = TempValue1 + x
                    Value1 = TempValue1
                    Current_key = Earring_Parameters_Dict[key]
                if Earring_Parameters_Dict[key] == "PT":
                    Value1 = Value1.upper()
                    Current_key = Earring_Parameters_Dict[key]
                else:
                    Current_key = Earring_Parameters_Dict[key]
                    StructuredPair = Earring_Parameters_Dict[key] + ":" + Value1

                Current_Item_Parameters_Dict[
                    Current_key
                ] = Value1  # list of sanitized parameters for the current item
                ITEM_OUTPUT_STRING = (
                    ITEM_OUTPUT_STRING
                    + "<"
                    + Earring_Parameters_Dict[key]
                    + ">"
                    + Value1
                    + "</"
                    + Earring_Parameters_Dict[key]
                    + ">"
                )
    else:
        LOCATION = NextLocation(LOCATION)
        LOCATION_String = "<LOCATION>" + LOCATION + "</LOCATION>"
        SKU = year_string + month_string + day_string + LOCATION + "000"
        SKU_String = "<SKU>" + SKU + "</SKU>"
        if PRE_PROCESSOR:
            print(
                f"Results of LOCATION Calculations: {LOCATION},{LOCATION_String}",
                {SKU},
                {SKU_String},
            )

        URLs = []
        PATHs = []
        URL_STRING = ""
        PATH_STRING = "<Pictures>"
        temp_URL = ""
        temp_PATH = ""
        current_image_counter = 0
        current_SixBit_image_counter = 0
        for i in range(1, PHOTOS_PER_ITEM + 1):
            current_image_counter = 5 * ItemCounter - i
            temp_URL = (
                base_url
                + "20"
                + year_string
                + "/"
                + month_string
                + "/"
                + day_string
                + "/"
                + batch_number
                + "/"
                + "IMG_"
                + imagenumberstrings[current_image_counter]
                + ".JPG"
            )
            if i == 5:
                current_SixBit_image_counter = current_image_counter + 4
                temp_PATH = (
                    base_path
                    + "\\20"
                    + year_string
                    + "\\"
                    + "100-"
                    + month_string
                    + day_string
                    + "-"
                    + batch_number
                    + "\\"
                    + "IMG_"
                    + imagenumberstrings[current_SixBit_image_counter]
                    + ".JPG"
                )
            elif i == 4:
                current_SixBit_image_counter = current_image_counter + 2
                temp_PATH = (
                    base_path
                    + "\\20"
                    + year_string
                    + "\\"
                    + "100-"
                    + month_string
                    + day_string
                    + "-"
                    + batch_number
                    + "\\"
                    + "IMG_"
                    + imagenumberstrings[current_SixBit_image_counter]
                    + ".JPG"
                )
            elif i == 3:
                temp_PATH = (
                    base_path
                    + "\\20"
                    + year_string
                    + "\\"
                    + "100-"
                    + month_string
                    + day_string
                    + "-"
                    + batch_number
                    + "\\"
                    + "IMG_"
                    + imagenumberstrings[current_image_counter]
                    + ".JPG"
                )
            elif i == 2:
                current_SixBit_image_counter = current_image_counter - 2
                temp_PATH = (
                    base_path
                    + "\\20"
                    + year_string
                    + "\\"
                    + "100-"
                    + month_string
                    + day_string
                    + "-"
                    + batch_number
                    + "\\"
                    + "IMG_"
                    + imagenumberstrings[current_SixBit_image_counter]
                    + ".JPG"
                )
            elif i == 1:
                current_SixBit_image_counter = current_image_counter - 4
                temp_PATH = (
                    base_path
                    + "\\20"
                    + year_string
                    + "\\"
                    + "100-"
                    + month_string
                    + day_string
                    + "-"
                    + batch_number
                    + "\\"
                    + "IMG_"
                    + imagenumberstrings[current_SixBit_image_counter]
                    + ".JPG"
                )
            else:
                print("Error creating SixBit .jpg URLs")

            URLs.append(temp_URL)
            PATHs.append(temp_PATH)
            URL_STRING = (
                URL_STRING
                + "<photoURL"
                + str(i)
                + ">"
                + temp_URL
                + "</photoURL"
                + str(i)
                + ">"
            )
            PATH_STRING = (
                PATH_STRING + "<Picture><Path>" + temp_PATH + "</Path></Picture>"
            )
        PATH_STRING = PATH_STRING + "</Pictures>"

        OUTPUT_STRING = (
            OUTPUT_STRING
            + "<Earring>"
            + SKU_String
            + LOCATION_String
            + ITEM_OUTPUT_STRING
            + "<Quantity>1</Quantity>"
            + URL_STRING
            + "</Earring>"
            + "\n"
        )

        #   *******  MAIN BODY OF 6Bit XML (Title, Height, Width):
        if Current_Item_Parameters_Dict["Title"]:
            TitleTagString = str(Current_Item_Parameters_Dict.get("Title"))
        else:
            TitleTagString = "Default Title"
        #               PRINT TO LOG FILE
        TitleTagString = TitleTagString.strip()
        TitleTag = "<Title>" + TitleTagString + "</Title>"
        if Current_Item_Parameters_Dict["Height"]:
            HeightTag = (
                "<DimensionLength>"
                + str(Current_Item_Parameters_Dict.get("Height"))
                + "</DimensionLength>"
            )
        else:
            HeightTag = "<DimensionLength>0.75</DimensionLength>"
        #            PRINT TO LOG FILE
        if Current_Item_Parameters_Dict["Width"]:
            WidthTag = (
                "<DimensionWidth>"
                + str(Current_Item_Parameters_Dict.get("Width"))
                + "</DimensionWidth>"
            )
        else:
            WidthTag = "<DimensionWidth>0.375</DimensionWidth>"
        #            PRINT TO LOG FILE
        EnabledOnEtsy = "<EnabledOnEtsy>true</EnabledOnEtsy>"
        Main_XML_Section = "<Item>" + TitleTag + HeightTag + WidthTag + EnabledOnEtsy

        #   *******  Etsy section of 6Bit XML (ShopSectionID, ShopSectionName, Tags, Materials, Description )
        ShippingPresetID = "<ShippingPresetID>150268745492</ShippingPresetID>"
        ShippingPresetName = "<ShippingPresetName>DomesticEarings</ShippingPresetName>"
        if Current_Item_Parameters_Dict["Type"]:
            Website_Type = str(Current_Item_Parameters_Dict.get("Type"))
            Website_Type = Website_Type.title()
            Website_Type = Website_Type.strip()
            SixBitSectionName = Section_Names_By_Type_Dict[Website_Type]
            SixBitSectionName = (
                "<ShopSectionName>" + SixBitSectionName + "</ShopSectionName>"
            )
        else:
            Website_Type = "French Hook"
            SixBitSectionName = Section_Names_By_Type_Dict[Website_Type]
            SixBitSectionName = (
                "<ShopSectionName>" + SixBitSectionName + "</ShopSectionName>"
            )

        # ShopSectionID =  need to complete list of ShopSectionIDs

        t = ""
        ConcatenatedTagsTemp = ""
        ConcatenatedTags = ""
        for t in TagsParametersList:
            ConcatenatedTagsTemp = ""
            ConcatenatedTagsTemp = Current_Item_Parameters_Dict[t] + ";"
            ConcatenatedTagsTemp = ConcatenatedTagsTemp[0:18]
            if ConcatenatedTagsTemp != ";":
                if ConcatenatedTagsTemp not in ConcatenatedTags:
                    ConcatenatedTags = ConcatenatedTags + ConcatenatedTagsTemp.strip()
        ConcatenatedTags = ConcatenatedTags[:-1]
        EtsyTags = "<Tags>" + ConcatenatedTags + "</Tags>"

        m = ""
        for m in MaterialsParametersList:
            ConcatenatedMaterialsTemp = ""
            ConcatenatedMaterialsTemp = Current_Item_Parameters_Dict[m] + ";"
            ConcatenatedMaterialsTemp = ConcatenatedMaterialsTemp[0:20]
            if ConcatenatedMaterialsTemp != ";":
                if ConcatenatedMaterialsTemp not in ConcatenatedMaterials:
                    ConcatenatedMaterials = (
                        ConcatenatedMaterials + ConcatenatedMaterialsTemp.strip()
                    )
        if ConcatenatedMaterials != "":
            ConcatenatedMaterials = ConcatenatedMaterials[:-1]
        else:
            ConcatenatedMaterials = "default material"
        EtsyMaterials = "<Materials>" + ConcatenatedMaterials + "</Materials>"

        CategoryIDElement = Section_IDs_By_Type_Dict[Website_Type]
        CategoryID = "<CategoryID>" + CategoryIDElement + "</CategoryID>"
        CategoryName = (
            "<CategoryName>"
            + Section_Names_By_Type_Dict[Website_Type]
            + "</CategoryName>"
        )

        ProcessingMin = "<ProcessingMin>1</ProcessingMin>"
        ProcessingMax = "<ProcessingMax>1</ProcessingMax>"
        WhoMade = "<WhoMade>someone_else</WhoMade>"
        WhatIsIt = "<WhatIsIt>FinishedProduct</WhatIsIt>"
        WhenMade = "<WhenMade>wm_before_2002</WhenMade>"

        EtsyDescriptionTemplate = "-- Selected for Etsy from our collection of thousands of vintage earrings at MyEarringDepot.com!"
        EtsyFooter = "  Your 5 star review is our best advertising and allows us to continue to offer thousands of vintage earrings at affordable prices.\
                      Your complete satisfaction is our goal.  If you are satisfied please give us a positive review.  If we have failed to \
                      satisfy you in any way, won't you please contact us immediately.  We'd greatly appreciate the opportunity to please you.  "
        Etsy_Description = (
            "<Description>"
            + str(Current_Item_Parameters_Dict.get("Description"))
            + EtsyDescriptionTemplate
            + EtsyFooter
            + "</Description>"
        )
        Etsy_Section = (
            "<Etsy>"
            + TitleTag
            + ShippingPresetID
            + ShippingPresetName
            + SixBitSectionName
            + EtsyTags
            + EtsyMaterials
            + CategoryID
            + CategoryName
            + ProcessingMin
            + ProcessingMax
            + WhoMade
            + WhenMade
            + Etsy_Description
            + "</Etsy>"
        )

        #   *******  Variation section of 6Bit XML (SKU, Quantity, Price)
        Variation_SKU = SKU_String
        if Current_Item_Parameters_Dict["Quantity"] == "":
            Variation_QtyToList = "<QtyToList>" + "1" + "</QtyToList>"
        else:
            Variation_QtyToList = (
                "<QtyToList>"
                + str(Current_Item_Parameters_Dict.get("Quantity"))
                + "</QtyToList>"
            )
        if Current_Item_Parameters_Dict["Price"] == "":
            Variation_FixedPriceEtsy = "<FixedPriceEtsy>14.99</FixedPriceEtsy>"
        else:
            Variation_FixedPriceEtsy = (
                "<FixedPriceEtsy>"
                + str(Current_Item_Parameters_Dict.get("Price"))
                + "</FixedPriceEtsy>"
            )
        Variation_Section = (
            "<Variations><Variation>"
            + Variation_SKU
            + Variation_QtyToList
            + Variation_FixedPriceEtsy
            + "</Variation></Variations>\n"
        )

        SixBitOutputString = (
            SixBitOutputString
            + Main_XML_Section
            + Etsy_Section
            + Variation_Section
            + PATH_STRING
            + "</Item>"
        )

        Current_Item_Parameters_Dict = {
            "Price": "",
            "PriceTier": "",
            "Quantity": "",
            "Title": "",
            "Description": "",
            "ShortDescription": "",
            "Type": "Post Dangle",
            "Color1": "",
            "Color2": "",
            "Metal1": "",
            "Metal2": "",
            "Gemstone1": "",
            "Gemstone2": "",
            "Material1": "",
            "Material2": "",
            "Mineral1": "",
            "Mineral2": "",
            "Height": "",
            "Width": "",
            "Diameter": "",
            "Era": "",
            "era": "",
            "Condition": "",
            "Style": "",
            "Wearer1": "",
            "Wearer2": "",
            "Theme1": "",
            "Theme2": "",
            "Occasion1": "",
            "Occasion2": "",
            "Character1": "",
            "Character2": "",
            "Brand1": "",
            "Brand2": "",
            "Shape1": "",
            "Shape2": "",
            "Tag1": "",
            "Tag2": "",
        }

        ItemCounter = ItemCounter + 1
        CurrentItemTermCount = 0
        ITEM_OUTPUT_STRING = ""
        ConcatenatedTags = ""
        ConcatenatedMaterials = ""
        ConcatenatedEtsyDescription = ""
    CumulativeTermCount = CumulativeTermCount + 1  # moves us through the list


# **********************END MAIN LOOP **********************************************************

OUTPUT_STRING = (
    XML_HEADER
    + "\n"
    + XML_COPYRIGHT_VERSION_NOTICE
    + "\n"
    + OPENING_PRODUCT_TAG
    + "\n"
    + OUTPUT_STRING
    + "</Product>"
)
SixbitOutputString = (
    XML_COPYRIGHT_VERSION_NOTICE
    + "\n"
    + SixBitOpeningTags
    + "\n"
    + SixBitOutputString
    + "\n"
    + SixBitClosingTags
)

xmloutputfile = "c:\\users\\rogeridaho\\" + batch_date + batch_number + ".xml"
sixbitxmloutputfile = (
    "c:\\users\\rogeridaho\\" + "6Bit" + batch_date + batch_number + ".xml"
)
if PRE_PROCESSOR:
    print(f"MED XML output file is:  {xmloutputfile}")
    print(f"(Etsy XML filename is: {sixbitxmloutputfile}")
    time.sleep(DEBUG_WAIT)

if LIVE:
    with open(xmloutputfile, "w", encoding="utf-8") as f:
        f.write(OUTPUT_STRING)

if PRE_PROCESSOR:
    dom = parseString(OUTPUT_STRING)
    print(dom.toprettyxml())

if LIVE:
    with open(sixbitxmloutputfile, "w", encoding="utf-8") as f:
        f.write(SixbitOutputString)

if PRE_PROCESSOR:
    dom = parseString(SixbitOutputString)
    print(dom.toprettyxml())

if LIVE:
    print(
        f" {current_image_counter} Photos processed, XML files created, program complete"
    )
