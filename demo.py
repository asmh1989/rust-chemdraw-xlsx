from lxml import etree
from bs4 import BeautifulSoup
from pathlib import Path
import pandas as pd


def parseSheetXML(xmlfile):
    file = open(xmlfile, "r")
    contents = file.read()
    # parsing
    soup = BeautifulSoup(contents, "xml")
    oleObjects = soup.find_all("oleObject")
    res_dict = {"rid": [], "row": [], "col": []}
    for ioleObj in oleObjects:
        # ioleObj=oleObjects[0]
        iattr = ioleObj.attrs
        if iattr["progId"] != "ChemDraw.Document.6.0":
            continue
        ioleObj = etree.HTML(str(ioleObj))
        try:
            col = ioleObj.xpath("//from/*[local-name()='col']")[0].text
            row = ioleObj.xpath("//from/*[local-name()='row']")[0].text
            res_dict["col"].append(col)
            res_dict["row"].append(row)
            res_dict["rid"].append(iattr["r:id"])
        except:
            continue
    return res_dict


def parseRelationXML(xmlfile):
    file = open(xmlfile, "r")
    contents = file.read()
    # parsing
    soup = BeautifulSoup(contents, "xml")
    oleObjects = soup.find_all("Relationship")
    res_dict = {"rid": [], "fileId": []}
    for ioleObj in oleObjects:
        # ioleObj=oleObjects[0]
        iattr = ioleObj.attrs
        if (
            iattr["Type"]
            != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
        ):
            continue
        fileId = Path(iattr["Target"]).name
        res_dict["rid"].append(iattr["Id"])
        res_dict["fileId"].append(fileId)
    return res_dict


import olefile
from zipfile import ZipFile
from glob import glob
from pathlib import Path
import os, sys
import subprocess
import pandas as pd


workDir = "demo"
filename = "demo.xlsx"
zip = ZipFile(filename, "r")
# print(zip.infolist())
workPath = Path(workDir)
workPath.mkdir(exist_ok=True, parents=True)
""" Write Sheet1.xml  """
f = zip.open("xl/worksheets/sheet1.xml").readlines()
with open(workPath.joinpath("sheet1.xml"), "wb") as output_file:
    for line in f:
        output_file.write(line)
f = zip.open("xl/worksheets/_rels/sheet1.xml.rels").readlines()
with open(workPath.joinpath("sheet1_relation.xml"), "wb") as output_file:
    for line in f:
        output_file.write(line)

entryPath = workPath.joinpath("cdxFiles")
entryPath.mkdir(exist_ok=True, parents=True)
smi_Dict = {"fileId": [], "smi": []}
for entry in zip.infolist():
    if not entry.filename.startswith("xl/embeddings/"):
        #   print(entry)
        continue

    f = zip.open(entry.filename)
    if not olefile.isOleFile(f):
        continue
    ole = olefile.OleFileIO(f)
    # print(ole.root.clsid)
    if ole.root.clsid != "41BA6D21-A02E-11CE-8FD9-0020AFD1F20C":  ## ChemDraw file type
        continue
    if not ole.exists("CONTENTS"):
        continue
    cdx_data = ole.openstream("CONTENTS").read()
    # print(cdx_data)
    cdxName = Path(entry.filename).name
    cdxFile = entryPath.joinpath(f"{cdxName}.cdx")
    with open(cdxFile, "wb") as output_file:
        output_file.write(cdx_data)

    # smi=os.system(f"obabel -icdx {cdxFile} -osmi")
    smi = subprocess.getoutput(f"obabel -icdx {cdxFile.as_posix()} -osmi")
    smi = smi.split()[0]
    # print(f"smi={smi}")
    smi_Dict["fileId"].append(cdxName)
    smi_Dict["smi"].append(smi)
dfSmi = pd.DataFrame.from_dict(smi_Dict)
# print(dfSmi)

xmlFile = f"{workDir}/sheet1.xml"
sheetDict = parseSheetXML(xmlFile)
dfSheet = pd.DataFrame.from_dict(sheetDict)
xmlFile = f"{workDir}/sheet1_relation.xml"
rlDict = parseRelationXML(xmlFile)
dfRl = pd.DataFrame.from_dict(rlDict)

dfRl = dfRl.merge(dfSheet, how="inner", on="rid")
dfSmi = dfSmi.merge(dfRl, how="inner", on="fileId")
print(dfSmi)
