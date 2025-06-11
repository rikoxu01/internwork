# _*_ coding  : UFT-8_*_
# 2020/4/3 13:55
# PyCharm

# Aim: download real-time announcement files in the SH stock exchange website

import pandas as pd
import requests
import re
import json
import datetime
import time
from time import sleep
import urllib.request
import os
from envelopes import Envelope
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import shutil
from lxml import etree

# genenrate a dataframe
def sheet_gen(search_sheet_text):
    output_sheet = pd.DataFrame(columns=['notice_name', 'date', 'url'])
    np_text = search_sheet_text[13:]
    #  search in np_text
    while len(np_text) > 5:
        np_date = np_text[1:11]
        np_nude_1 = np_text.find('</span><a href="')
        np_nude_2 = np_text.find('" target="_blank">')
        np_nude_3 = np_text.find('</a></dd>')
        np_url = np_text[(np_nude_1+16):np_nude_2]
        np_title = np_text[(np_nude_2+18):np_nude_3]
        add_row_i = pd.DataFrame([[np_title, np_date, np_url]], columns=output_sheet.columns)
        output_sheet = pd.concat([output_sheet, add_row_i], axis=0, sort=None)
        np_text = np_text[(np_nude_3+18):]
    output_sheet.reset_index(drop=True, inplace=True)
    return output_sheet


def sh_sheet():
    nowDate_str = time.strftime("%Y-%m-%d", time.localtime())
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    driver = webdriver.Chrome(options=chrome_options)
    driver.get('http://www.sse.com.cn/disclosure/bond/announcement/asset/')
    time.sleep(1)
    # import date, search dataframe
    from_date_input = driver.find_element_by_css_selector("#start_date")
    driver.execute_script("arguments[0].removeAttribute('readonly')", from_date_input)
    from_date_input.clear()
    from_date_input.send_keys(nowDate_str)
    from_date_input.send_keys(Keys.ESCAPE)
    sleep(1)
    to_date_input = driver.find_element_by_css_selector("#end_date")
    driver.execute_script("arguments[0].removeAttribute('readonly')", to_date_input)
    to_date_input.clear()
    to_date_input.send_keys(nowDate_str)
    to_date_input.send_keys(Keys.ENTER)
    sleep(1)
    to_date_input.send_keys(Keys.ESCAPE)
    # read the sheet
    url_list = []
    search_sheet = driver.find_element_by_css_selector("html.js.canvas.rgba.multiplebgs.cssanimations.csstransforms.csstransforms3d.csstransitions body div.page_content.bgimg1 div.container div.row div.col-sm-9 div.row div.col-sm-12 div.con_block div.sse_common_wrap_cn div.sse_wrap_cn_con div.sse_list_1")
    search_sheet_text = search_sheet.get_attribute('innerHTML')
    if search_sheet_text != '暂无数据' and search_sheet_text != '<dl><dd>暂无数据</dd></dl>':
        output_sheet = sheet_gen(search_sheet_text)
        np_button_list = driver.find_elements_by_css_selector("html.js.canvas.rgba.multiplebgs.cssanimations.csstransforms.csstransforms3d.csstransitions body div.page_content.bgimg1 div.container div.row div.col-sm-9 div.row div.col-sm-12 div.con_block div.sse_common_wrap_cn div.sse_wrap_cn_con div.page-con-table.js_pageTable nav.table-page.page-con.hidden-xs ul.pagination li a#idStr.classPage")
        while len(np_button_list) > 0:
            np_button_list[0].click()
            sleep(2)
            search_sheet = driver.find_element_by_css_selector(
                "html.js.canvas.rgba.multiplebgs.cssanimations.csstransforms.csstransforms3d.csstransitions body div.page_content.bgimg1 div.container div.row div.col-sm-9 div.row div.col-sm-12 div.con_block div.sse_common_wrap_cn div.sse_wrap_cn_con div.sse_list_1")
            search_sheet_text = search_sheet.text
            output_sheet_i = sheet_gen(search_sheet_text)
            output_sheet = pd.concat([output_sheet, output_sheet_i], axis=o, sort=None)
            output_sheet.reset_index(inplace=True, drop=True)
    else:
        output_sheet = pd.DataFrame(columns=['notice_name', 'date', 'url'])
    driver.quit()
    return output_sheet


def send_mail_a(to_addr,subject,text_body, attachment_addr1):
  from_addr = r"username@xyz.com"
  password = r"password"
  smtp_server = "smtp.exmail.qq.com"
  envelope = Envelope(from_addr=from_addr,to_addr=to_addr,subject=subject,text_body=text_body)
  envelope.add_attachment(attachment_addr1)
  #envelope.add_attachment(attachment_addr2)
  #envelope.add_attachment(attachment_addr3)
  envelope.send(smtp_server,login=from_addr, password=password, tls=True)


def send_mail_b(to_addr,subject,text_body):
  from_addr = r"username@xyz.com"
  password = r"password"
  smtp_server = "smtp.exmail.qq.com"
  envelope = Envelope(from_addr=from_addr,to_addr=to_addr,subject=subject,text_body=text_body)
  envelope.send(smtp_server,login=from_addr, password=password, tls=True)


def file_download(download_folder, file_name, download_url):
    file_name = file_name.replace(':', '：')
    list_number = 0
    folder_list = os.listdir(download_folder)
    for name_i in folder_list:
        if name_i.find(file_name) != -1:
            list_number = list_number + 1
    if list_number == 0:
        download_path = download_folder + file_name + '\\'
    else:
        download_path = download_folder + file_name + str(list_number) + '\\'
    os.makedirs(download_path)
    chrome_options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': download_path}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--headless')
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(download_url)
    time.sleep(6)
    driver.quit()
    os.rename(download_path + os.listdir(download_path)[0], download_path + file_name + '.pdf')
    np_url = download_path + file_name + '.pdf'
    return np_url


mail_list = [
    "name1@xyz.com",
    "name2@xyz.com",
]


if __name__ == '__main__':
    while True:
        # This code is only run between 6AM and 8PM (Beijing time)
        if 6 < datetime.datetime.now().hour < 20: 
            print("start")
            print(datetime.datetime.now())
            try:
                nowDate = time.strftime("%Y%m%d", time.localtime())
                nowDate_str = time.strftime("%Y-%m-%d", time.localtime())
                np_path = "D:\\ABS Files\\Exchange_Data\\SSE\\SSE" + nowDate + ".xlsx"
                download_folder = "D:\\ABS Files\\Exchange_Data\\SSE\\SSE\\" + nowDate + "\\"
                if os.path.exists(np_path):
                    or_sheet_sh = pd.read_excel(np_path)
                    sheet_sh = sh_sheet()
                    if sheet_sh.shape[0] > 0:
                        send_mail_times = 0
                        for index, row in sheet_sh.iterrows():
                            if row['url'] not in or_sheet_sh['url'].tolist():
                                # down load files
                                np_add = file_download(download_folder, row['notice_name'], row['url'])
                                # send emails to teammates, and save in Excel
                                for mail_i in mail_list:
                                    send_mail_a(mail_i, row['notice_name'], row['date'] + '-上交所\n' + row['url'], np_add)
                                or_sheet_sh = or_sheet_sh.append(row, ignore_index=True)
                                send_mail_times = send_mail_times + 1
                                #  send_mail_a('xxx@xxx.com', row['newstitle'], 'Best wishes!')
                        if send_mail_times >= 1:
                            print("send mail ready")
                            or_sheet_sh.to_excel(np_path, index=False)
                else:
                    sheet_sh = sh_sheet()
                    if sheet_sh.shape[0] > 0:
                        sheet_sh.to_excel(np_path, index=False)
                        os.makedirs(download_folder)
                        for index, row in sheet_sh.iterrows():
                            if row['date'] == nowDate_str:
                                #  down load files
                                try:
                                    np_add = file_download(download_folder, row['notice_name'], row['url'])
                                    # send emails
                                    for mail_i in mail_list:
                                        send_mail_a(mail_i, row['notice_name'], row['date'] + '-上交所\n' + row['url'], np_add)
                                except:
                                    for mail_i in mail_list:
                                        send_mail_b(mail_i, row['notice_name'], row['date'] + '-上交所\n' + row['url'])
                            print("send mail ready")
                print("success")
                print("")
                sleep(120)
            except:
                print('failed')
                sleep(120)
        else:
            sleep(120)