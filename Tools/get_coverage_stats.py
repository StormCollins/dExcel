"""
This script parses the code coverage report to get the code coverage percentage.

Specifically, during the "test" job of the CI pipeline, OpenCover generates "results.xml".
ReportGenerator then converts "results.xml" to an HTML report, "Report/index.html".
And finally this script searches for the "Line Coverage" and prints it to the GitLab console.
pipeline job during CI.

In order for the script to work the following software needs to be installed on Mabi (the runner).
    • Python
    • bs4 (BeautifulSoup) - a Python package,
    • lxml - a Python package,
"""

from bs4 import BeautifulSoup
from lxml import etree


try:
    report_path: str = r'publish\CoverageResults\Report\index.html'
    f: FileIo = open(report_path, 'r')
    page: str = ''.join(f.readlines())
    soup = BeautifulSoup(page, features='lxml')
    dom = etree.HTML(str(soup))
    coverage: str = dom.xpath(r"//html/body/div[1]/div/div[1]/div[2]/div[2]/div[2]/table/tr[5]/td/text()")[0]
    print(f'Coverage: {coverage}')
except Exception as err:
    print(f"Unexpected {err=}, {type(err)=}")
    raise