import requests
import sys


def requests_download(url, filename):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36"}
    content = requests.get(url, headers = headers).content
    with open(filename, 'wb') as file:
        file.write(content)


def get_url(filename):
    f = open(filename, "r+", encoding='utf8')
    data = f.read()
    pos2 = data.find("100w_100h")
    pos2 = pos2 - 1
    pos1 = data.find("image")
    pos1 = pos1 + 18
    url_str = data[pos1: pos2]
    return url_str


if __name__ == "__main__":
    url = sys.argv[1]
    tempDic = sys.argv[2]

    filename = tempDic + "\\temp.htm"
    picname = tempDic + "\\temp_pic_path"

    requests_download(url, filename)
    url_str = get_url(filename)

    url_str = "https://" + url_str
    requests_download(url_str, picname)
