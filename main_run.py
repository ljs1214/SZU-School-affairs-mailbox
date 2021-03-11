import requests

import re

import collections

import xlwt


re_mod = re.compile('\d{4}-.{3,5}')  # 正则匹配日期
date = []
for i in range(1, 156):  # 总共125页建议信箱
    url = "https://www1.szu.edu.cn/mailbox/list.asp?page="+str(i)+"&leader=%BD%A8%D1%D4%CF%D7%B2%DF&tag="  # 用page确定页数
    r = requests.get(url)
    r_text = r.text
    date += (re_mod.findall(r_text))
    print(str((i/156)*100)[:4]+"%")  # 进度条
for i in range(1, 271):  # 总共223页投诉信箱
    url = "https://www1.szu.edu.cn/mailbox/list.asp?page="+str(i)+"&leader=%CE%CA%CC%E2%CD%B6%CB%DF&tag="  # 用page确定页数
    r = requests.get(url)
    r_text = r.text
    date += (re_mod.findall(r_text))
    print(str((i/271)*100)[:4]+"%")  # 进度条
date_dict = dict(collections.Counter(date))
print(date_dict)  # 至此爬取完毕，下面是输出为Excel


def excel_output():
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('My Worksheet')
    date_list_key = list(date_dict.keys())
    date_list_value = list(date_dict.values())
    for i in range(len(date_list_key)):
        worksheet.write(i, 0, date_list_key[i])
        worksheet.write(i, 1, date_list_value[i])
    workbook.save('Excel_test.xls')


excel_output()
