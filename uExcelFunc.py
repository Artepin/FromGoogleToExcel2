import openpyexcel

def findCell(wb, p_id):
    for i in wb.sheetnames:
        #print(i)
        sheet = wb[i]
        namedRanges = str(wb.defined_names[p_id].attr_text).split('$')
        namedRanges = namedRanges[1]+namedRanges[2]
        #print(namedRanges)
        if namedRanges !=None:
            cell = sheet[namedRanges]
            return cell