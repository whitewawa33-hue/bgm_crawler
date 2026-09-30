# -*- coding = utf-8 -*-
# @Time:2025/4/15
# @Name:hutao
# @Title:ai pc
# -*- coding: utf-8 -*-
import re
import time
import urllib.request
from bs4 import BeautifulSoup
import xlwt
import logging
import os


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

#主流程函数
def main():
    start = int(input("请输入起始页: "))
    end = int(input("请输入结束页: "))
    file_name = (input("请输入保存文件名:"))+'.xls'
    base_url = 'https://bgm.tv/anime/browser?sort=rank&page=' #网站url前缀
    datalist = []
    logger.info(f"Start, page range: {start} to {end}")

    for page in range(start-1, end): #页码范围（第300-301页）
        current_url = f"{base_url}{page + 1}"
        html = ask_url(current_url) #调用函数，获取网址
        if html:
            datalist.extend(get_data(html)[1:-1]) #调用函数，获取数据
        time.sleep(0.5)
        logger.info(f"Scrape page {page+1}")

    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    file_path = os.path.join(output_dir, file_name)
    save_data(datalist, file_path) #调用函数，数据保存为xls文件

    logger.info(f"Total items collected: {len(datalist)}")

def ask_url(url):
    #防反爬，网站无需登陆，不需要cookie
    headers = {
        'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Mobile Safari/537.36'
    }
    req = urllib.request.Request(url=url, headers=headers, method='GET')

    try:
        with urllib.request.urlopen(req) as response:
            return response.read().decode('utf-8')
    except urllib.error.URLError as e:
        logger.error(f"Error fetching {url}: {e}")
        return None


def get_data(html):
    soup = BeautifulSoup(html, 'html.parser')
    results = []

    for item in soup.find_all('div', class_="inner"):
        item_str = str(item)
        entry = [
            re.findall(r'<a class="l" href="/subject/.*?">(.*?)</a>', item_str),  # 名字
            re.findall(r'</small> <span class="tip_j">\((.*?)\)</span>', item_str),  # 评论数
            re.findall(r'<small class="fade">(.*?)</small>', item_str)  # 评分
        ]
        results.append(entry)

    return results


def save_data(data_list, filename):
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('reviews', cell_overwrite_ok=True)

    # 写入表头
    headers = ('名字', '评论数', '评分')
    for col, header in enumerate(headers):
        sheet.write(0, col, header)

    # 写入数据
    for row, entry in enumerate(data_list, start=1):
        for col, value in enumerate(entry):
            sheet.write(row, col, value[0] if value else '')  # 取列表第一个元素或空字符串

    book.save(filename)
    logger.info(f"Data saved to {filename}")


if __name__ == '__main__':
    main()