import time
import xlwt
import xlrd
from selenium import webdriver

"""
    使用说明：
    该文件运行依赖于火狐浏览器驱动，下载地址 https://pan.baidu.com/s/1kC1SSjgOD1Xm8Lt4vAX4GA
    将 geckodriver.exe 文件放置与改文件同一目录下，若无法使用，请设置环境变量,或是检查驱动版本与软件是否版本相同！
    安装火狐浏览器
    运行该程序即可

    “嘻嘻”单词导入整理使用
    根据日语单词生成对应假名，以便导入单词卡片
    将单词写入【单词.xls】文件  ----》 生成后将在【单词生成.xls】中
                                                    by 满脑子电线的程序员
"""


def other_query(driver,send_keys):
    time.sleep(1)
    xpath = driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[1]/form/div[1]/input')
    xpath.clear()
    xpath.send_keys(send_keys)
    print("输入了" + send_keys)
    time.sleep(1)
    try:
        driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[1]/form/div[2]/div[2]/button[1]').click()
    except:
        pass
    kana = driver.find_element_by_xpath(
        '/html/body/div[1]/div/main/div/section/div/section/div/header/div[1]/div[1]/h2').text
    print(kana)
    japanese = driver.find_element_by_xpath(
        '/html/body/div[1]/div/main/div/section/div/section/div/header/div[1]/div[2]/span[1]').text
    print(japanese[1:-1])
    return send_keys ,kana, japanese[1:-1]

def first_query(driver,send_keys):
    xpath = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/form/div[1]/input')
    xpath.send_keys(send_keys)
    print("输入了"+send_keys)
    time.sleep(1)
    try:
        driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/form/div[2]/div[2]/button[1]').click()
    except:
        pass
    kana = driver.find_element_by_xpath(
        '/html/body/div[1]/div/main/div/section/div/section/div/header/div[1]/div[1]/h2').text
    print(kana)
    japanese = driver.find_element_by_xpath(
        '/html/body/div[1]/div/main/div/section/div/section/div/header/div[1]/div[2]/span[1]').text
    print(japanese[1:-1])
    return send_keys ,kana, japanese[1:-1]

def init_Brower():
    driver = webdriver.Firefox()

    driver.set_page_load_timeout(3)
    driver.set_script_timeout(3)  # 这两种设置都进行才有效
    try:
        driver.get('https://dict.hjenglish.com/jp')
    except:
        pass
    time.sleep(1)
    return driver

# 通过读取xml获得单词
def read_excel():
    data = xlrd.open_workbook('单词.xls')
    table = data.sheet_by_index(0) #通过索引顺序获取
    values = table.col_values(0)
    print(values)
    return values

def write_excel_close(data):
    data.save("单词生成.xls")

def write_excel_init():
    data = xlwt.Workbook()  # 新建一个excel
    sheet = data.add_sheet('单词生成')  # 添加一个sheet页
    return data,sheet

def write_excel(send_keys, kana, japanese,number,sheet):
    sheet.write(number, 0, send_keys)
    sheet.write(number, 1, kana)
    sheet.write(number, 2, japanese)

if __name__ == '__main__':
    number = 0
    xml = read_excel()
    data, sheet = write_excel_init()
    brower = init_Brower()
    send_keys, kana, japanese = first_query(brower, xml[0])
    write_excel("原汉字", "假名", "汉字", number, sheet)
    number += 2
    write_excel(send_keys,japanese, kana, number, sheet)
    number += 1
    xml.remove(xml[0])
    for i in xml:
        send_keys, kana, japanese = other_query(brower,i)
        write_excel(send_keys, japanese, kana,number,sheet)
        number +=1
    print(send_keys, kana, japanese)

    write_excel_close(data)
