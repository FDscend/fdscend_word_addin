import requests
import sys


def requests_download(url, filename):
    content = requests.get(url).content
    with open(filename, 'wb') as file:
        file.write(content)


def get_url(filename):
    f = open(filename, "r+", encoding='utf8')
    data = f.read().encode('utf-8').decode("unicode_escape")
    pos1 = data.find("\"cover\"") + 9
    pos2 = data.find("\"tags\"") - 2
    url_str = data[pos1: pos2]
    return url_str


if __name__ == "__main__":
    url = sys.argv[1]
    tempDic = sys.argv[2]

    filename = tempDic + "\\temp.htm"
    picname = tempDic + "\\temp_pic_path"

    requests_download(url, filename)
    url_str = get_url(filename)

    requests_download(url_str, picname)

