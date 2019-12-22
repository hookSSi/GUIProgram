from calendar import monthrange
from datetime import datetime
from bs4 import BeautifulSoup
import datefinder
import requests
import math
import os

# 만들어진 폴더 이름 반환
def CreateDir(dir_name):
    dir_path = os.path.join(dir_name)

    try:
        if not(os.path.isdir(dir_name)):
            os.makedirs(dir_path)
    except OSError as e:
        if e.errno != errno.EEXIST:
            print("Failed to create directory!!!!!")
            raise

    return dir_path

def download(url, file_name, path = ".", isLogin = False):
    
    if(isLogin):
        cookies = {}
        headers = {}

        LOGIN_INFO = {
                        'id':'sounghoo4699', 'pw' : '4699sounghoo'
                     }

        with requests.Session() as s:
            login_req = s.post("http://academy.myvilpt.gethompy.com/manager/admin/admin_login.php", data = LOGIN_INFO)
            cookies = s.cookies
            headers = s.headers
            with open(path + "/" + file_name, "wb") as file:
                response = s.get(url)
                file.write(response.content)
    else:
        with open(path + "/" + file_name, "wb") as file:
            response = requests.get(url)
            file.write(response.content)

    return path + "/" + file_name

# 중복 제거
def erase_overlap(args):
    result = set(args)
    result = list(result)
    return result

def extract_date(s):
    date = list(datefinder.find_dates(text = s))[0]
    temp = datetime(date.year, date.month, date.day, 0, 0)
    return temp

def get_last_day(year, month):
    last_day = monthrange(year, month)[1]
    return last_day

def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def string_colnum(s):
    num = 0

    length = len(s)
    ch_range = 26
    for i in range(0, length):
        temp = math.pow(ch_range, i) + (ord(s[(length - 1) - i]) - 65)
        num += int(temp)

    return num

# [시,군,구]를 리턴함
def SearchArea(keyword):
        url =   "http://postcode.map.daum.net/search?region_name={0}".format(keyword) \
              + "&cq=&cpage=1&origin=http%3A%2F%2Fpostcode.map.daum.net&isp=N&isgr=N&isgj=N&ongr=&ongj=&regionid=&regionname=&roadcode=&roadname=&banner=on&indaum=off&vt=layer&am=on&ani=on&mode=view&sd=on&hmb=off&heb=off&asea=off&smh=off&zo=on&theme=&bit=&sit=&sgit=&sbit=&pit=&mit=&lcit=&plrg=&plrgt=1.5&us=on&msi=10&ahs=off&whas=500&zn=Y&sm=on&CWinWidth=1903&sptype=&sporgq=&fullpath=%2Fguide&a51=off"
        r = requests.get(url)
        soup = BeautifulSoup(r.text, 'html.parser')
        selector = soup.find("select",{'id':'inpArea'})
        area_name = selector.find("option",{'data-idx':'1'})

        if area_name is None:
            print("주소가 잘못되었습니다.")
            return None
        else:
            area_name = area_name.text.split(' ')
            return [area_name[0], area_name[1]]

def CheckValidData(data, check_data_list, check):
    for e in check_data_list:
        if(check(data)):
            return True
    return False

# 딕셔너리 디버그
def DicDebug(dic):
    for key in dic.keys():
        print(key)
        print(dic[key])

def replace_all(text, dic):
    for i, j in dic.items():
        text = text.replace(i, j)
    return text