from bs4 import BeautifulSoup
from lxml import etree
import os

print(f'Current Working Directory: {os.getcwd()}')
print('')
print(f'Files/folders in current directory:')
for file in os.scandir('publish/CoverageResults/Report'):
    print(file)

print('')
report_path: str = r'publish\CoverageResults\Report\index.html'
f = open(report_path, 'r')
page = ''.join(f.readlines())
soup = BeautifulSoup(page, features='lxml')
dom = etree.HTML(str(soup))
print(f'Coverage: {dom.xpath(r"//html/body/div[1]/div/div[1]/div[2]/div[2]/div[1]/text()")[0]}')
