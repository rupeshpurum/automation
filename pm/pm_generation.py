# Generating pm xml as per the atributes mentioned in the xlsx file.

__author__='Rupesh'
__status__ = 'Prototype'

import time
start = time.time()
import os
import fileinput
from glob import glob

import pandas as pd
import xml.etree.ElementTree as ET
from pandas import ExcelFile
xl = pd.ExcelFile('example_modified.xlsx')

print(xl.sheet_names)
#print(len(xl.sheet_names))


for sheet in xl.sheet_names:
    print(sheet)
    a = sheet

# sh refers to sheet

    sh = pd.read_excel('example_modified.xlsx',sheet)
    for index, row in sh.iterrows():
        print(row['Montype'], row['Units'])
        # a = row['Units']
        # t = type('12')
        # for idx,i in enumerate(a):
        #     if type(i) != t:
        #         print(idx,type(i))


        root = ET.Element("metricGroup", id=sheet+" Statistics.TL1", name=sheet+" Statistics TL1", protocol="TL1", displayType="Normal")
        # for i in range(mt):
        for index, row in sh.iterrows():

            doc = ET.SubElement(root, "metric", id=sheet + " "+row['Montype'] + "." + row["Location"] + " " + row["Direction"],
                                name=sheet + " " + row['Montype'] + " " + row["Location"] + " " + row["Direction"],
                                desc=row['Metric'],
                                protocol="TL1",
                                units=' ' if row['Units'] is None else row['Units'],
                                conversion_function="PER_PERIOD",
                                consolidation_function="SUM",
                                location='NEAR_END' if row['Location'] == 'NEND' else "FAR_END" if row['Location'] == 'FEND' else 'NA',
                                direction='TRANSMIT' if ((row['Direction'] == 'Tx') or (row['Direction'] == 'TX')or (row['Direction'] == 'TRMT')) else 'RECEIVE' if row['Direction'] == ('Rx' or 'RX' or 'RCV') else 'NA',
                                min="0",
                                displayType="lineSeries-hist" if ((row['Type'] == 'gauge') or (row['Type'] == 'GAUGE')) else "verticalBar-hist",
                                displayColor="BLUE" if ((row['Direction'] == 'Tx') or (row['Direction'] == 'TX') or (row['Direction'] == 'TRMT')) else "DARK_GREEN")


            ET.SubElement(doc,"parameter",name=row["Montype"],
                          collector="TL1",
                          type='COUNTER' if ((row['Type'] == 'counter') or (row['Type'] == 'COUNTER') or (row['Type'] == 'Counter')) else 'GAUGE' if ((row['Type'] == 'gauge') or (row['Type'] == 'GAUGE') or (row['Type'] == 'Gauge')) else 'NA',
                          oid=row["Montype"])

            ET.SubElement(doc, "value", parameter=row["Montype"])

            tree = ET.ElementTree(root)
            # open(sheet, 'a').close()
            # print(tree.getroot())
            print(row['Montype'],type(row['Montype']))
            cwd = os.getcwd()
            os.chdir("/home/rupesh/TL1/pm/")
            tree.write(a + "_tl1_new.xml")
            os.chdir(cwd)


os.chdir("/home/rupesh/TL1/pm/")

import glob
import errno
path = '/home/rupesh/TL1/pm/*.xml'
files = glob.glob(path)
for name in files:
    try:
        with open(name) as f:
            filedata = f.read()
        filedata = filedata.replace('</metric>', '\n\t</metric>')
        filedata = filedata.replace('consolidation_function','\n\t\tconsolidation-function')
        filedata = filedata.replace('conversion_function','\n\t\tconversion-function')
        filedata = filedata.replace("desc=\"","\n\t\tdesc=\"")

        filedata = filedata.replace("direction=\"","\n\t\tdirection=\"")

        filedata = filedata.replace("displayColor","\n\t\tdisplayColor")

        filedata = filedata.replace("\" displayType=\"","\"\n\t\tdisplayType=\"")
        filedata = filedata.replace("hist\" id=\"","hist\" \n\t\tid=\"")
        filedata = filedata.replace("location=\"","\n\t\tlocation=\"")
        filedata = filedata.replace("\" min=\"0\" ","\"\n\t\tmin=\"0\"\n\t\t")
        filedata = filedata.replace("protocol=\"TL1\" units=\"", "\n\t\tprotocol=\"TL1\" \n\t\tunits=\"")
        filedata = filedata.replace("<parameter","\n\n\t\t <parameter")
        filedata = filedata.replace("><",">\n\t<")
        filedata = filedata.replace("<metricGroup","\t<metricGroup")

        filedata = filedata.replace("<value parameter","\t <value parameter")
        filedata = filedata.replace("<metric","\t<metric")
        filedata = filedata.replace("\t\t<metric","\t<metric")
        filedata = filedata.replace("</metricGroup","\t</metricGroup")
        filedata = filedata.replace("\t<metricGroup","  <metricGroup")
        filedata = filedata.replace("\t</metricGroup","  </metricGroup")
        filedata = filedata.replace("\t  </metricGroup>","  </metricGroup>")
        filedata = filedata.replace(">\n\n\t\t <","\n\t >\n\n\t\t <")

        filedata = filedata.replace("<metricGroup","\n\n<metricGroup")
        # NA is renamed as No in sheet to overcome exception.
        #
        #
        # Hence it has to be renamed after processing the data.
        #
        # Below line to to do the same
        filedata = filedata.replace(" No\""," NA\"")


        with open(name, 'w') as f:
            f.write(filedata)
    except IOError as exc:
        if exc.errno != errno.EISDIR:
            raise

print(os.getcwd())

# Handling error:  TypeErro: argument of type  'float' is not iterable
# ib = df_test ["income_bracket"]
# t = type('12')
# for idx,i in enumerate(ib):
#     if(type(i) != t):
#         print idx,type(i)


end = time.time()
print('Time taken for the pm generation in seconds:',end-start)
