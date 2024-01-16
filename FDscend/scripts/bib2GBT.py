import json
import sys


def getAuthor(bibJson):
    _author = bibJson["author"].replace(' and ',",").replace(' & ',",").upper()
    
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
    fileBGT = getAuthor(bibJson) + ". " + bibJson["title"] + reftype[bibtype] + ". " + bibJson["school"] + "," + bibJson["year"]
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

    fileBGT = (getAuthor(bibJson) + ". " + 
               bibJson["title"] + 
               reftype[bibtype] + ". " + 
               bibJson["journal"] +
               _year + _volume + _number + _pages
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
               _year + _pages
               )

    return fileBGT


def type_N(bibJson, reftype, bibtype):
    
    fileBGT = (getAuthor(bibJson) + ". " + 
               bibJson["title"] + 
               reftype[bibtype] + ". " + 
               bibJson["institution"] + "," + 
               bibJson["year"]
               )

    return fileBGT


def mainProscess(bibFile: str) -> str:
    with open(bibFile, "r", encoding='utf-8') as f:
        lines = f.readlines() 


    index_1 = lines[0].find('@')
    index_2 = lines[0].find('{')
    bibtype = lines[0][index_1 + 1: index_2].replace(' ', '')


    jsonLines =[]
    jsonLines.append('{\r\n')

    for l in lines[1: -1]:
        if (l.find("=") != -1) & (l.find('{') == -1) & (l.find('}') == -1):
            index_1 = l.find("=")
            index_2 = l.rfind(",")
            l = '"' + l[0: index_1].replace(' ', '') + '":"' + l[index_1 + 1: index_2].strip() + '",\r\n'
        else:
            l = l.replace('{', '"').replace('}', '"')
            index = l.find("=")

            if(index != -1): l = '"' + l[0: index].replace(' ', '') + '":' + l[index + 1: -1] + "\r\n"

        jsonLines.append(l)

    jsonLines.append(lines[-1])

    index = jsonLines[-2].rfind('"')
    jsonLines[-2] = jsonLines[-2][0: index] + jsonLines[-2][index: -1].replace(',', '') + "\r\n"

    bibJson = json.loads("".join(jsonLines).replace('\r', '').replace('\n', ''))
    reftype = {"article": "[J]",
                "mastersthesis": "[D]",
                "inproceedings": "[C]",
                "techreport": "[N]",
                "misc": "[P]",
                "manual": "[P]"}

    match reftype[bibtype]:
        case "[J]":
            fileBGT = type_J(bibJson, reftype, bibtype)
        case "[D]":
            fileBGT = type_D(bibJson, reftype, bibtype)
        case "[C]":
            fileBGT = type_C(bibJson, reftype, bibtype)
        case "[N]":
            fileBGT = type_N(bibJson, reftype, bibtype)
        case _:
            fileBGT = "** 类型不支持 **"
   
    return fileBGT


if __name__ == "__main__":
    bibFile = sys.argv[1]
    # bibFile = input("bibtex file:\r\n")

    print(mainProscess(bibFile))
