from xlsx import workbook

#Note that, the date in xlsx must be string, not date format

visaType = ['F1', 'F2A', 'F2B', 'F3', 'F4', 'EB1', 'EB2', 'EB3','EB_OW']
countryType = ['all', 'China-mainland-born', 'INDIA', 'MEXICO', 'PHILIPPINES']
visaTypeRow = {'F1':'2','F2A':'3','F2B':'4','F3':'5','F4':'6','EB1':'7','EB2':'8','EB3':'9','EB_OW':'10'};
countryColumn = {'all':'B','China-mainland-born':'C','INDIA':'D','MEXICO':'E','PHILIPPINES':'F'};

def process():
    for vType in visaType:
        for c in countryType:
            print '%s-%s.json'% (vType, c)
            generateData(vType, c, countryColumn[c] + visaTypeRow[vType])

def generateData (vType, cType, position):
    vFile = open('file_list.txt');
    fileString = vFile.read()
    vFile.close()
    fileList = fileString.split();

    output = "{"
    output += "\"type\":\"" + vType +"\","
    output += "\"country\":\"" + cType +"\","

    output += "\"dateArray\":["

    for i in fileList:
        name = i.split('.')
        date = name[0].split('_')
        d = "%s/%s" % (date[1] ,date[0])
        actionDate = getDateAt(position, i);
        output += '{\"publishedDate\":\"%s\",\"actionDate\":\"%s\"},' % (d, actionDate)

    output = output[:-1] # remove the last character
    output += "]}"

    outputFileName = vType + "_" + cType + ".json"
    outFile = open(outputFileName, "w")
    outFile.write (output)
    outFile.close()


def getDateAt (position, fileName):
    Workbook = workbook(fileName)
    sheet = Workbook.Sheets.visa
    cell = sheet[position]
    return cell

if __name__ == '__main__':
    process()
