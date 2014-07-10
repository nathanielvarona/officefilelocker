#!/usr/bin/env jython
# Requires Jython and Apache POI
import os, sys, getopt, ConfigParser, mimetypes

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
from org.apache.poi.hssf.usermodel import *
from org.apache.poi.hwpf.usermodel import *
from org.apache.poi.hwpf import HWPFDocument
from org.apache.poi.xwpf.usermodel import *
# from org.apache.poi.poifs.crypt import HashAlgorithm
from java.io import FileInputStream
from java.io import FileOutputStream

xssf_supported_mimetypes = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
]

hssf_supported_mimetypes = [
    "application/vnd.ms-excel"
]

xwpf_supported_mimetypes = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
]

hwpf_supported_mimetypes = [
    "application/msword"
]


def main(argv):
    username = 'user'
    password = ''
    input_file = ''
    output_file = ''

    def get_mimetype(input_file):
        file_mimetype, file_mime_encoding = mimetypes.guess_type(input_file)
        return file_mimetype, file_mime_encoding

    def usage():
        print '\nUsage: jython %s -u <username> -p <password> -i <inputfile> -o <outputfile>\n' % sys.argv[0]
        sys.exit()

    try:
        opts, args = getopt.getopt(argv,"hu:p:i:o:",["username=","password=","input=","output="])
    except getopt.GetoptError:
        usage()
    for opt, arg in opts:
        if opt == '-h':
            usage()
        elif opt in ("-u", "--username"):
            username = arg
        elif opt in ("-p", "--password"):
            password = arg
        elif opt in ("-i", "--input"):
            input_file = arg
        elif opt in ("-o", "--output"):
            output_file = arg

    if password and input_file and output_file:
        mimetype = get_mimetype(input_file)[0]
        # print "Office File MimeType: %s" % mimetype
        fileIn = FileInputStream(input_file)
        fileOut = FileOutputStream(output_file)

        if mimetype in xssf_supported_mimetypes:
            workbook = XSSFWorkbook(fileIn)
            worksheets = workbook.getNumberOfSheets()
            for worksheet in range(0,worksheets):
                sheet = workbook.getSheetAt(worksheet)
                sheet.protectSheet(password)
                sheet.enableLocking();
            workbook.lockStructure();
            workbook.write(fileOut)

        elif mimetype in hssf_supported_mimetypes:
            workbook = HSSFWorkbook(fileIn)
            workbook.writeProtectWorkbook(password, username)
            workbook.write(fileOut)

        elif mimetype in xwpf_supported_mimetypes:
            worddocument = XWPFDocument(fileIn)
            # ToDo: Protect the Document with Password
            # http://poi.apache.org/apidocs/org/apache/poi/xwpf/usermodel/XWPFDocument.html#enforceReadonlyProtection(java.lang.String, org.apache.poi.poifs.crypt.HashAlgorithm)
            worddocument.enforceReadonlyProtection()
            worddocument.write(fileOut)

        elif mimetype in hwpf_supported_mimetypes:
            # ToDo: Convert HWPFDocument to XWPFDocument File Stream
            # http://poi.apache.org/apidocs/org/apache/poi/hwpf/HWPFDocument.html
            # worddocument = HWPFDocument(fileIn)
            # print dir(worddocument)
            pass

        fileIn.close()

    else:
        usage()

if __name__ == "__main__":
    main(sys.argv[1:])
