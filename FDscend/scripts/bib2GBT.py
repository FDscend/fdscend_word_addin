import json
import sys
import re


def getAuthor(bibJson):
    if bibJson["author"].find('&') != -1:  # cnki
        _author = bibJson["author"].replace(' and ',",").replace(' & ',",").replace('.',"").upper()
    else:
        _author = bibJson["author"].replace(',',"").replace(' and ',",").replace('.',"").upper()

    _author = _author.replace(', ',",").replace(',',", ")

    _author = re.sub(r"\{[^{}]*\}", "", _author)
    
    if(_author.count(",") >= 3):
        index_1 = _author.find(',')
        index_2 = _author[index_1 + 1: -1].find(',')
        index_3 = _author[index_1 + 1: -1][index_2 + 1: -1].find(',')

        etc = ", et al"

        if('\u4e00' <= _author[0: index_1] <= '\u9fff'):
            etc = ",等"
        
        _author = _author[0: index_1 + index_2 + index_3 + 2] + etc

    return _author


def type_D(bibJson, reftype, bibtype):
    fileBGT = getAuthor(bibJson) + ". " + bibJson["title"] + reftype[bibtype] + ". " + bibJson["school"] + "," + bibJson["year"] + "."
    return fileBGT


def type_J(bibJson, reftype, bibtype):
    
    if "year" in bibJson: _year = "," + bibJson["year"] + ','
    else: _year = ""

    if "volume" in bibJson: _volume = bibJson["volume"]
    else: _volume = ""

    if "number" in bibJson: _number = '(' + bibJson["number"] + ")"
    else: _number = ""

    if "pages" in bibJson: _pages = ':' + bibJson["pages"]
    else: _pages = ""

    # cnki
    if ("number" in bibJson) & ("pages" not in bibJson):
        _number = ""
        _pages = ':' + bibJson["number"]

    fileBGT = (getAuthor(bibJson) + ". " + 
               bibJson["title"] + 
               reftype[bibtype] + ". " + 
               bibJson["journal"] +
               _year + _volume + _number + _pages + "."
               )

    return fileBGT


def type_C(bibJson, reftype, bibtype):
    
    if "year" in bibJson: _year = "," + bibJson["year"] + ','
    else: _year = ""

    if "pages" in bibJson: _pages = ':' + bibJson["pages"]
    else: _pages = ""

    fileBGT = (getAuthor(bibJson) + ". " + 
               bibJson["title"] + 
               reftype[bibtype] + ". //" + 
               bibJson["booktitle"] +
               _year + _pages + "."
               )

    return fileBGT


def type_N(bibJson, reftype, bibtype):
    
    fileBGT = (getAuthor(bibJson) + ". " + 
               bibJson["title"] + 
               reftype[bibtype] + ". " + 
               bibJson["institution"] + "," + 
               bibJson["year"] + "."
               )

    return fileBGT


def getBibJson(lines: list) -> dict:

    jsonLines =[]
    jsonLines.append('{\r\n')


    for l in lines[1: -1]:
        if (l.find('"') != -1) & (l.find('"') != l.rfind('"')):
            index_1 = l.find("=")
            index_2 = l.find('"')
            index_3 = l.rfind('"')
            l = '"' + l[0: index_1].replace(' ', '') + '":' + l[index_2: index_3] + '",\r\n'   

        else:
            #cnki
            if ((l.find('=') != -1) &
                (l.find('"') == -1) &
                (l.find('{') == -1) &
                (l.find('}') == -1)):

                index_1 = l.find("=")
                index_2 = l.find(',')
                l = '"' + l[0: index_1].replace(' ', '') + '":"' + l[index_1 + 1: index_2].strip() + '",\r\n'
            
            else:
                l = l.replace('{', '"', 1)[::-1].replace('}', '"', 1)[::-1]  # replace the outermost layer only
                index = l.find("=")
                if(index != -1): l = '"' + l[0: index].replace(' ', '') + '":' + l[index + 1: -1] + "\r\n"

        jsonLines.append(l)

    jsonLines.append('}')

    index = jsonLines[-2].rfind('"')
    jsonLines[-2] = jsonLines[-2][0: index] + jsonLines[-2][index: -1].replace(',', '') + "\r\n"
    
    bibJson = json.loads("".join(jsonLines).replace('\r', '').replace('\n', '').replace('\t', '').replace('\\', ''))

    return bibJson


def mainProscess(bibFile: str) -> str:
    with open(bibFile, "r", encoding='utf-8') as f:
        lines = f.readlines() 

    lines_multibib = [[]]
    lines_singlebib = []
    bib_count = 0

    for l in lines:
        if l == "\n": continue
        if l.find('@') != -1:
            if bib_count % 2 == 1:
                lines_multibib.insert(bib_count - 1, lines_singlebib[:])
                lines_singlebib.clear()
            bib_count += 1
    
        lines_singlebib.append(l)

    lines_multibib.insert(bib_count - 1, lines_singlebib[:])
    lines_multibib.pop(bib_count)


    reftype = {"article": "[J]",
            "mastersthesis": "[D]",
            "phdthesis": "[D]",
            "inproceedings": "[C]",
            "conference": "[C]",
            "book": "[M]",
            "booklet": "[M]",
            "techreport": "[N]",
            "misc": "[P]",
            "manual": "[P]"}


    fileBGT = []

    for _lines in lines_multibib:
   
        index_1 = _lines[0].find('@')
        index_2 = _lines[0].find('{')
    
        if index_2 == -1: index_2 = _lines[0].find('(')
    
        bibtype = _lines[0][index_1 + 1: index_2].replace(' ', '').lower()
    
        # print(bibtype)
    
        index_3 = _lines[-1].rfind(')')
    
        if index_3 != -1:
            _lines[-1] = _lines[-1][0: index_3]
            _lines.append('}')
        
        bibJson = getBibJson(_lines)
    
        match reftype[bibtype]:
            case "[J]":
                fileBGT.append(type_J(bibJson, reftype, bibtype))
            case "[D]":
                fileBGT.append(type_D(bibJson, reftype, bibtype))
            case "[C]":
                fileBGT.append(type_C(bibJson, reftype, bibtype))
            case "[N]":
                fileBGT.append(type_N(bibJson, reftype, bibtype))
            case _:
                fileBGT.append("** 类型不支持 **")

   
    return fileBGT


if __name__ == "__main__":
    bibFile = sys.argv[1]
    # bibFile = input("bibtex file:\r\n")

    fileBGT = mainProscess(bibFile)
    _count = 1
    for l in fileBGT:
        if fileBGT.__len__() == 1: print(l)
        else:
            print(f'[{_count}]', l)
            _count += 1
