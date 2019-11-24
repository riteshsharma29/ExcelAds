#!/usr/bin/python

import sys,codecs
import xlrd

reload(sys)  # Reload does the trick!
sys.setdefaultencoding('UTF8')

'''if len(sys.argv) < 2:
    print ""
    print ""
    print "Error !!! Please pass first parameter as ads filename"
    print ""
    print ""
    sys.exit()

adsfile = sys.argv[1]'''

adsfile = "CX_eX2_QML_ADS_2.9.xls"

langlist = ["english","t. chinese","s. chinese","japanese","korean","french","german"]

#this list is defined to handle logic of last value for a particular language of the sheet
decmntval = [7,6,5,4,3,2,1]


book = xlrd.open_workbook(adsfile)

#stores sheet names
#print book.sheet_names()[1]


def genlabels(langvar,decval,jsonfile):

    jsnfile = codecs.open(jsonfile, "w", "utf-8")
    jsnfile.write("{" + chr(10) + chr(10))

    count = 0

    for sht,sheets in enumerate(book.sheets()):

        lastsheet = len(book.sheets())
        lastsheet = lastsheet - 1

    #for col in range(sheets.ncols):
        if sht > 0:
            lastrow = len(range(sheets.nrows))
            for rows in range(3,lastrow):

                #append sheetname with opening brace
                if rows == 3 :jsnfile.write("    " +  chr(34) + book.sheet_names()[sht] + chr(34) + ":{" + chr(10))

                if sht > 0 :
                    #reading language [numric value after rows, - is the column number/index]
                    lang = (str(sheets.cell(rows, 1).value)).strip().lower()

                    #reading description
                    desc = (str(sheets.cell(rows, 2).value)).strip()
                    desc = desc.encode('utf-8')
                    desc = desc.replace(chr(10),"<br>").replace(".0","")

                    # calculating the last row for specfic language based on decval var
                    if rows != lastrow - decval:
                        desc = chr(34) + desc + chr(34) + ","
                    elif rows == lastrow - decval and sht != lastsheet:
                        desc = chr(34) + desc + chr(34) + chr(10) + "   },"
                    elif rows == lastrow - decval and sht == lastsheet:
                        desc = chr(34) + desc + chr(34) + chr(10) + "   }" + chr(10) + "}"

                    #reading tag
                    if lang == "english":tag = (str(sheets.cell(rows, 5).value)).strip()

                    #reading instance
                    inst = (str(sheets.cell(rows, 6).value)).strip().lower()

                    #language 3 letter extension
                    langext = lang[0:3]
                    langext = langext.replace("jap","jpn").replace("t.","tch").replace("s.","sch")

                    if inst == "all":
                        tagval = chr(34) + tag + "_" + langext + chr(34) + ":"
                        genstr = "               " + tagval + desc + chr(10)
                        genstr = genstr.decode('utf-8')
                        jsnfile.write(genstr)
                        #print genstr
                    elif lang == langvar:
                        tagval = chr(34) + tag + chr(34) + ":"
                        genstr = "               " + tagval + desc + chr(10)
                        genstr = genstr.decode('utf-8')
                        jsnfile.write(genstr)
                        #print genstr

genlabels(langlist[0],decmntval[0],"Labels_eng.jsn")
genlabels(langlist[1],decmntval[1],"Labels_tch.jsn")
genlabels(langlist[2],decmntval[2],"Labels_sch.jsn")
genlabels(langlist[3],decmntval[3],"Labels_jpn.jsn")
genlabels(langlist[4],decmntval[4],"Labels_kor.jsn")
genlabels(langlist[5],decmntval[5],"Labels_fre.jsn")
genlabels(langlist[6],decmntval[6],"Labels_ger.jsn")



'''for sheet in book.sheets():
    for row in range(sheet.nrows):
        #print sheet.row(row)
        print row'''