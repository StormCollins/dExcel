from bs4 import BeautifulSoup
from lxml import etree
import os

print(f'Current Working Directory: {os.getcwd()}')
report_path: str = r'publish\CoverageResults\Report\index.html'
f = open(report_path, 'r')
page = ''.join(f.readlines())
soup = BeautifulSoup(page, features='lxml')
dom = etree.HTML(str(soup))
coverage = dom.xpath(r"//html/body/div[1]/div/div[1]/div[2]/div[2]/div[2]/table/tr[5]/td/text()")[0]
print(f'Coverage: {coverage}')