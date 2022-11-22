import logging
import os
import sys
from time import sleep
import win32com.client


# Logger config
logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
file_handler = logging.FileHandler('D:\SYSTEM\RTF_CONVERT_DOCX\RtfConvertor.log')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)


# Converting Function
def ConvertRtfToDocs(full_path):
    try:
        logger.info("Starting the convertation of " + full_path)
        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        wdFormatDocumentDefault = 16
        doc = word.Documents.Open(full_path)
        for pic in doc.InlineShapes:
            try:
                pic.LinkFormat.SavePictureWithDocument = True
            except:
                pass
        doc.SaveAs(str(full_path.split(".")[0] + ".docx"), FileFormat=wdFormatDocumentDefault)
        doc.Close
        word.Quit()
        logger.info("Converted {} successfully".format(full_path))
        return (True)
    except:
        logger.error("Unexpected error while trying to convert {} file".format(full_path))
        return (False)


# Directory Scanner

def scanner(root_dir):
    for path, dirs, files in os.walk(root_dir):
        for file in files:
            full_path = path + "\\" + file
            if file.split(".")[-1] == "rtf":
                if ConvertRtfToDocs(full_path):
                    os.remove(full_path)
                    logger.info("Original RTF file is deleted succefully")


# BU dirs to handle (scan and convert)
root_BU1 = r"C:\SystemRTFoutput1"
root_BU2 = r"C:\SystemRTFoutput2"
root_BU3 = r"C:\SystemRTFoutput3"
root_BU_all = [root_BU1, root_BU2, root_BU3]  # List of BUs to handle

flag_path = r"D:\SYSTEM\RTF_CONVERT_DOCX\Flag\flag.txt"  # flag_path

while True:
    sleep(1)
    if os.path.exists(flag_path):
        for BU in root_BU_all:
            scanner(BU)
        os.remove(flag_path)
        for BU in root_BU_all:
            scanner(BU)

