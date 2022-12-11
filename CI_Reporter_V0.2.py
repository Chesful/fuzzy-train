import pandas as pd
import datetime as dt
import numpy as np
import xlsxwriter

class ColorBackground:
    cianColour=['background-color: powderblue','background-color: #DAEEF3']
    orangeColour=['background-color: #FABF8F','background-color: #FDE9D9']
    outputColour=orangeColour
    def __init__(self, colourPalet=cianColour):
        self.colourPalet = colourPalet
        self.state=0
    
    def changeColourPalet(self):
        if self.colourPalet==ColorBackground.cianColour:
            self.colourPalet=ColorBackground.orangeColour
        else:
            self.colourPalet=ColorBackground.cianColour
    def changeState(self):
        if self.state==0:
            self.state=1
        else:
            self.state=0

def colourRows(row):
    global prevScriptName
    seriesList=[]
    if row['SCRIPT']!=prevScriptName:
        ColorRow.changeColourPalet()
        prevScriptName=row['SCRIPT']
    
    ColorRow.changeState()
    for cell in row:
        seriesList.append(ColorRow.colourPalet[ColorRow.state])
    return seriesList

def highlight_greaterthan(s, threshold=0, column=0):
    is_max = pd.Series(data=False, index=s.index)
    return ['background-color: yellow' for v in is_max]

def adjustCols(worksheet1):
    worksheet1.set_column(0, 0, 55)
    worksheet1.set_column(1, 1, 14)
    worksheet1.set_column(2, 2, 10)
    worksheet1.set_column(3, 3, 100)
    worksheet1.set_column(4, 4, 150)

filepath = "./Reports/report2022-12-11.xlsx"
scriptStheet = "Scripts"
dataScript = pd.read_excel(filepath, sheet_name=scriptStheet, usecols="A,D:G")
ColorRow=ColorBackground()

print(dataScript['RESULT'].describe())

today = dt.date.today().strftime("%Y-%m-%d")

LaunchScript = dataScript[dataScript.RESULT != "NOT_LAUNCHED"]
LaunchScript['DATE'] = pd.to_datetime(LaunchScript['DATE'], format="%d/%m/%Y %H:%M:%S").dt.strftime("%d/%m/%Y")
FailScripts = LaunchScript[LaunchScript.RESULT == 'FAIL']
FailScripts = FailScripts.dropna()

GroupFail = FailScripts.INFO.str.extract("\[(.*?)\]")
FailScripts.insert(5,"LINE",GroupFail[0])
FileScriptName = FailScripts.SCRIPT.drop_duplicates()

resultDF= pd.DataFrame({"SCRIPT":[],"DATE":[],"ELEMENT":[],"RESULT":[],"INFO":[],"LINE":[]})
FailResultDF= pd.DataFrame({"SCRIPT":[],"DATE":[],"Nº ELEM":[],"ELEMENTS":[],"INFO":[]})

css_odd_rows = 'background-color: powderblue; color: black;'
prevScriptName=""




for scriptFile in FileScriptName:
    FailScriptResultDF= pd.DataFrame({"SCRIPT":[],"DATE":[],"Nº ELEM":[],"ELEMENTS":[],"INFO":[]})
    workingDataframe = FailScripts[FailScripts.SCRIPT ==scriptFile]
    errorScriptList = workingDataframe.LINE.drop_duplicates()
    
    for errorline in errorScriptList:
        elementsWithSameError=workingDataframe[workingDataframe.LINE==errorline]
        ElementListSameError = elementsWithSameError.ELEMENT.tolist()
        ElementStrList=""
        for elementName in ElementListSameError:
            if len(ElementStrList)==0:
                ElementStrList=elementName
            else:
                ElementStrList=ElementStrList+"; "+elementName
        FailScriptResultDF = pd.concat([FailScriptResultDF,pd.DataFrame({"SCRIPT":[scriptFile],"DATE":[elementsWithSameError.DATE.iat[0]],"Nº ELEM":[len(ElementListSameError)],"ELEMENTS":[ElementStrList],"INFO":[elementsWithSameError.INFO.iat[0]]})])
    FailResultDF= pd.concat([FailResultDF,FailScriptResultDF])

FailResultDF=FailResultDF.reset_index(drop=True)
FailScriptsNames= FailScripts.drop_duplicates(subset="LINE")
Depured=FailScripts.drop_duplicates(subset="SCRIPT")
Depured.drop("LINE", inplace=True, axis=1)


print(dataScript['RESULT'].describe())

with pd.ExcelWriter("./Results/Result "+str(today)+".xlsx") as writer:
    LaunchScript.style.set_properties(**{'background-color': 'yellow'}).to_excel(writer,sheet_name="TestResults", index=False)
    FailResultDF.style.apply(colourRows, axis=1).to_excel(writer, sheet_name='ScriptFail', index=False)
    workSheet=writer.sheets['ScriptFail']
    (max_row, max_col) = FailResultDF.shape
    column_settings = [{'header': column} for column in FailResultDF.columns]
    workSheet.add_table('A1:E'+str(len(FailResultDF)), {'columns': column_settings})
    adjustCols(workSheet)