import pandas
import re
from os import listdir
from os.path import isfile, join

def reverseTable(table):
    tablebis = list(map(list, zip(*table)))
    return tablebis

def createIndex():
    mypath = "output/tables/"
    onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
    indexfile = "output/tableindex.tex"
    lines = []
    for file in onlyfiles:
        lines.append("\\input{tables/"+file+"}\n")
    open(indexfile, "w").writelines(lines)
    
def getData():
    filename = "data/Master2018Data.xlsx"
    xl = pandas.ExcelFile(filename)
    sn = xl.sheet_names
    for s in sn:
        r = sheetToList(s)
        sheetToTex(r, s)

def getTexColumnList(size):
    c = ""
    for i in range(0,size):
        c+="L"
    return c

def getSheetInfo(sheetName, label):
    inifile = "config/"+sheetName+".ini"
    tableInfoFile = open(inifile, "r")
    for line in tableInfoFile:
        m = re.search(r""+label+": \"(.*)\"", line)
        if m:
            return m.group(1)
    return ""

def sheetToTex(sheetData, sheetName):
    datalist = []
    title = getSheetInfo(sheetName, "title")
    label = getSheetInfo(sheetName, "label")
    source = getSheetInfo(sheetName, "source")
    filename = getSheetInfo(sheetName, "filename")
    reverse = eval(getSheetInfo(sheetName, "reverse"))
    
    if reverse:
        sheetData = reverseTable(sheetData)
        
    datalist.append("\\begin{table}[h!]\n")
    datalist.append("\\caption{"+title+"}\\label{table:"+label+"}\n")
    
    id = 0
    for row in sheetData:
        if id==0:
            header = row
            datalist.append("\\begin{tabularx}{\\textwidth}{@{\\extracolsep{6pt}}"+getTexColumnList(len(header))+"@{}}\n")
            datalist.append("\\toprule\n")
            datalist.append(" & ".join(texHeader(header))+"\\\\\n")
            datalist.append(headerRule(texHeader(header)))
            
        else:
            datalist.append(" & ".join(texRow(row))+"\\\\\n")
            datalist.append("\\midrule\n")
        id+=1
    datalist.append("\\bottomrule\n")
    datalist.append("\\end{tabularx}\n")
    datalist.append("\\flushright Souce: "+source+"\n")
    datalist.append("\\end{table}\n")
    open("output/tables/"+filename+".tex", "w").writelines(datalist)
    # return datalist
    
def texRow(row):
    texRow = []
    for x in row:
        m = re.search("[0-9]", str(x))
        m2 = re.search("[^0-9]", str(x))
        if m and not m2:
            x = "\\SI{"+str(x)+"}{}"
        texRow.append(str(x))
     
    return texRow

def texHeader(row):
    texHeader = []
    i = 0
    n=0
    masterValue = ""
    for cell in row:
        previousValue = ""
        nextValue = ""
        currentValue = cell
        
        if len(row)>i+1:
            nextValue = row[i+1]
        
        if len(row)>i-1:
            previousValue = row[i-1]
            
        if "Unnamed: " in cell:
            if "Unnamed: " not in previousValue:
                n=2
                masterValue = previousValue
            else:
                n+=1

            if "Unnamed: " not in nextValue:
                texHeader.append("\multicolumn{"+str(n)+"}{c}{"+masterValue+"}")
        
        elif "Unnamed: " not in nextValue:
            texHeader.append(currentValue)
        i+=1
    
    return texHeader

def headerRule(headerList):
    p=0
    rule = ""
    for headerCell in headerList:
        m = re.search("\\\\multicolumn\\{([0-9]+)\\}", headerCell)
        if m:
            rule+="\\cmidrule(lr){"+str(p+1)+"-"+str(eval(m.group(1))+p)+"}"
            p+=eval(m.group(1))
            
        else:
            p+=1
    if rule:
        return rule+"\n"
    else:
        return "\\midrule\n"
        
def sheetToList(sheetname):
    filename = "data/Master2018Data.xlsx"
    xl = pandas.ExcelFile(filename)
    r = []
    df = xl.parse(sheetname)
    df = df.fillna("")
    
    rowlist = [""]+list(df.index.values)
    newlist = []
    for column in df:
        newlist.append([column]+list(df[column]))
    r = []
    r.append(rowlist)
    r = r+newlist
    
    return r

getData()
createIndex()