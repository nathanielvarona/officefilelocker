#!/usr/bin/env jython
# This requires Jython and Apache POI
import os, sys, getopt, ConfigParser

try:
    config = ConfigParser.ConfigParser()
    config.read(os.path.join(os.path.dirname(os.path.abspath(__file__)), "apache.cfg"))
    APACHE_POI_JAR_PATH = config.get("POI", "path")
except Exception:
    APACHE_POI_JAR_PATH = os.environ["APACHE_POI_JAR_PATH"]

for root, dirs, files in os.walk(APACHE_POI_JAR_PATH):
    for name in files:
        if name.endswith(".jar"):
            sys.path.append(os.path.join(root, name))

from org.apache.poi.xssf.usermodel import *
from java.io import FileInputStream
from java.io import FileOutputStream


def main(argv):
    password = ''
    input_file = ''
    output_file = ''

    def message():
        print '\nUsage: jython %s -p <password> -i <inputfile> -o <outputfile>\n' % sys.argv[0]
        sys.exit()

    try:
        opts, args = getopt.getopt(argv,"hp:i:o:",["password=","input=","output="])
    except getopt.GetoptError:
        message()
    for opt, arg in opts:
        if opt == '-h':
            message()
        elif opt in ("-p", "--password"):
            password = arg
        elif opt in ("-i", "--input"):
            input_file = arg
        elif opt in ("-o", "--output"):
            output_file = arg

    if password and input_file and output_file:
        fileIn = FileInputStream(input_file)
        fileOut = FileOutputStream(output_file)

        workbook = XSSFWorkbook(fileIn)
        worksheets = workbook.getNumberOfSheets()

        for worksheet in range(0,worksheets):
            sheet = workbook.getSheetAt(worksheet)
            # sheet.sheetName.title()
            sheet.lockDeleteColumns()
            sheet.lockDeleteRows()
            sheet.lockFormatCells()
            sheet.lockFormatColumns()
            sheet.lockFormatRows()
            sheet.lockInsertColumns()
            sheet.lockInsertRows()
            sheet.protectSheet(password)
            sheet.enableLocking();

        workbook.lockStructure();
        workbook.write(fileOut)
        fileIn.close()
    else:
        message()

if __name__ == "__main__":
    main(sys.argv[1:])
