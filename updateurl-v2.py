#!/usr/bin/python
# -*- coding: UTF-8 -*-
from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import numpy as np
import xlwings as xw
import requests
import pandas as pd
import time
import re

driver = webdriver.Chrome(ChromeDriverManager().install())


def delhtml(inhtml):
    html = inhtml
    dr = re.compile(r'<[^>]+>', re.S)
    dd = dr.sub('', html)
    # 去空格和多换行符
    refinedd = dd.replace(' ', '').replace('\n\n\n\n\n\n\n', '\n').replace('\n\n\n\n\n\n', '\n').replace(
        '\n\n\n\n\n', '\n').replace('\n\n\n\n', '\n').replace('\n\n\n', '\n').replace('\n\n', '\n')
    # print(refinedd)
    return refinedd
# Python通过正则表达式去除(过滤)HTML标签示例代码:
# 过滤HTML中的标签
# 将HTML中标签等信息去掉
# @param htmlstr HTML字符串.


def filter_tags(htmlstr):
    # 先过滤CDATA
    re_cdata = re.compile('//<!\[CDATA\[[^>]*//\]\]>', re.I)  # 匹配CDATA
    re_script = re.compile(
        '<\s*script[^>]*>[^<]*<\s*/\s*script\s*>', re.I)  # Script
    re_style = re.compile(
        '<\s*style[^>]*>[^<]*<\s*/\s*style\s*>', re.I)  # style
    re_br = re.compile('<br\s*?/?>')  # 处理换行
    re_h = re.compile('</?\w+[^>]*>')  # HTML标签
    re_comment = re.compile('<!--[^>]*-->')  # HTML注释
    s = re_cdata.sub('', htmlstr)  # 去掉CDATA
    s = re_script.sub('', s)  # 去掉SCRIPT
    s = re_style.sub('', s)  # 去掉style
    s = re_br.sub('\n', s)  # 将br转换为换行
    s = re_h.sub('', s)  # 去掉HTML 标签
    s = re_comment.sub('', s)  # 去掉HTML注释
    # 去掉多余的空行
    blank_line = re.compile('\n+')
    s = blank_line.sub('\n', s)
    s = replaceCharEntity(s)  # 替换实体
    return s
# 替换常用HTML字符实体.
# 使用正常的字符替换HTML中特殊的字符实体.
# 你可以添加新的实体字符到CHAR_ENTITIES中,处理更多HTML字符实体.
#@param htmlstr HTML字符串.


def replaceCharEntity(htmlstr):
    CHAR_ENTITIES = {'nbsp': ' ', '160': ' ',
                     'lt': '<', '60': '<',
                     'gt': '>', '62': '>',
                     'amp': '&', '38': '&',
                     'quot': '"', '34': '"', }
    re_charEntity = re.compile(r'&#?(?P<name>\w+);')
    sz = re_charEntity.search(htmlstr)
    while sz:
        entity = sz.group()  # entity全称，如>
        key = sz.group('name')  # 去除&;后entity,如>为gt
        try:
            htmlstr = re_charEntity.sub(CHAR_ENTITIES[key], htmlstr, 1)
            sz = re_charEntity.search(htmlstr)
        except KeyError:
            # 以空串代替
            htmlstr = re_charEntity.sub('', htmlstr, 1)
            sz = re_charEntity.search(htmlstr)
    return htmlstr

#=================================开始执行抓取代码
driver = webdriver.Chrome(ChromeDriverManager().install())
wb = xw.Book('queryfirm.xlsm')
sheet0 = wb.sheets[0]
all_row = sheet0.range('A1').expand().last_cell.row
#for v_1 in range(1, all_row - 10):
for v_1 in range(1, 3):
    content = sheet0[v_1, 0].value
    print(f'sheet[{v_1},0]={content}')
    # url='https://www.qixin.com/search?key=%E7%A6%8F%E5%BB%BA%E4%B8%AD%E5%A5%A5%E5%98%89%E4%BF%A1%E6%81%AF%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8&page=1'
    url = 'https://www.qixin.com/search?key=' + \
        sheet0[v_1, 109].value + 'page=1'
    # 根据公司名查启信宝
    driver.get(url)
    soup = BeautifulSoup(driver.page_source, 'lxml')
    # 检测是否有出现防“爬虫”的“点击按钮进行验证”，若有就等待
    # for bu in soup.button:
    #   #print(tt.string)
    #   if bu.find('点击按钮进行验证')>1:
    #     print(tt.string)
    #     #driver.find_element_by_xpath("/html/body/div[2]/div/div/div/div/button").click()
    #     ##等10秒
    #     print("Start : %s" % time.ctime())
    #     time.sleep( 10 )
    #     print("End   : %s" % time.ctime())
    #   else:
    #     print ('..........正常抓取\n')
    divs = soup.select('.col-2.clearfix')
    if (len(divs) <= 0):
        # 查无记录，等10秒
        print("等20秒\n")
        print("Start : %s" % time.ctime())
        time.sleep(5)
        print("End   : %s" % time.ctime())
        # 重新取
        # driver.find_element_by_xpath("/html/body/div[2]/div/div/div/div/button").click()
        driver.get(url)
        soup = BeautifulSoup(driver.page_source, 'lxml')
    # 根据类名选择
    divs = soup.select('.col-2.clearfix')
    href = divs[0].select('a')[0].get('href')
    allurl = 'https://www.qixin.com' + href
    sheet0[v_1, 111].value = allurl
    # 根据启信宝的信息查第二页数据
    # driver.get('https://www.qixin.com/company/695b78e7-46c4-4994-abcc-65e795be0a92')
    driver.get(allurl)
    content = driver.page_source
    # delcon=delhtml(content)
    sheet0[v_1, 112].value = filter_tags(content)
    # 根据启信宝的信息查第二页数据
    # 发起人股东
    sheet0[v_1, 113].value = driver.find_element_by_xpath(
        "/html/body/div[2]/div[6]/div/div[1]/div[1]/div/div[2]/table").text
    # 主要人员
    sheet0[v_1, 114].value = driver.find_element_by_xpath(
        "/html/body/div[2]/div[6]/div/div[1]/div[1]/div/div[3]").text
    # 变更记录
    sheet0[v_1, 115].value = driver.find_element_by_xpath(
        "/html/body/div[2]/div[6]/div/div[1]/div[1]/div/div[5]").text
    # 工商信息
    sheet0[v_1, 116].value = driver.find_element_by_xpath(
        "/html/body/div[2]/div[6]/div/div[1]/div[1]/div/div[1]").text

wb.save()
wb.close()


# for i in range(0,5):
#     print(i)
#sheet0 = wb.sheets[0]
# sheet0["A2"].value
# for div in divs:
#     href=div.select('a')[0].get('href')
#     allurl='https://www.qixin.com/'+href
#     print(allurl)
# 取第一条匹配的公司
