#!/usr/bin/python

import pandas as pd
import sys,codecs

'''if len(sys.argv) < 2:
    print ""
    print ""
    print "Error !!! Please pass first parameter as ads filename"
    print ""
    print ""
    sys.exit()

adsfile = sys.argv[1]'''

adsfile = "CX_eX2_QML_ADS_2.9.xls"

xlfile = pd.ExcelFile(adsfile)


langlist = ['English','T. Chinese','S. Chinese','Japanese','Korean','French','German']
langext = ['eng','tch','sch','jap','kor','fre','ger']

def genlabls(adsfile,sheetno,starow,l,lindx,wrkshtn):

    jsnfile = codecs.open(langext[lindx] + ".jsn", "a+", "utf-8")
    jsnfile.write("{" + chr(10) + chr(10))

    df = pd.read_excel(adsfile,sheetname=sheetno,header=starow)
    langs = df['Languages'].values
    lastrow = len(langs) - 1

    jsnfile.write("    " + chr(34) + wrkshtn + chr(34)  + ":{"  + chr(10))

    for row,lang in enumerate(langs):
        lang = str(lang)
        lang = lang.strip()

        inst = df['Instances'].values[row]

        if lang == "English":
            tag = df['Tag Name'].values[row]
            tag = str(tag)
            tag = tag.encode('utf-8')
            tag = tag.strip()

        if lang == l and str(inst).lower() != "all" :

            lang = lang.replace(". Ch","ch").replace("Jap","jpn")
            ext = lang[0:3]
            ext = ext.lower()

            desc = df['Description'].values[row]
            if type(desc) != int and type(desc) != float: desc = desc.encode('utf-8')
            desc = str(desc)
            desc = desc.strip()
            desc = " " + chr(34) + desc + chr(34)

            tagv = "               " + chr(34) + tag + chr(34) + ":"

            genstr = tagv + desc
            genstr = genstr.decode('utf-8')

            genstr = genstr + "," + chr(10)

            jsnfile.write(genstr)

        if lang in langlist and str(inst).lower() == "all" :

            lang = lang.replace(". Ch","ch").replace("Jap","jpn")
            ext = lang[0:3]
            ext = ext.lower()
            desc = df['Description'].values[row]
            if type(desc) != int and type(desc) != float:desc = desc.encode('utf-8')
            desc = str(desc)
            desc = desc.strip()
            desc = " " + chr(34) + desc + chr(34)

            tagv = "               " + chr(34) + tag + "_" + ext + chr(34) + ":"
            genstr = tagv + desc
            genstr = genstr.decode('utf-8')

            genstr = genstr + "," + chr(10)

            jsnfile.write(genstr)

    jsnfile.write("   }," + chr(10))


for sheetno,sheetname in enumerate(xlfile.sheet_names):
    if sheetno != 0:
        genlabls(adsfile, sheetno, 2, "English", 0, sheetname)
        genlabls(adsfile, sheetno, 2, "French", 5, sheetname)
