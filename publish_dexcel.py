# This script is used to publish the latest release build of dExcel to the shared drive.
import os
import re
import xml.etree.ElementTree as ET
import zipfile
import shutil


def zipdir(path, ziph):
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file),
                       os.path.relpath(os.path.join(root, file), path))


release_build_path: str = r'C:\GitLab\dExcelTools\dExcel\dExcel\bin\Release\net6.0-windows'
shared_drive_releases_path: str = r'\\ZAJNB010\Capital Markets 2\AQS Quants\dExcelTools\Releases'
dexcel_project_file_path: str = r'C:\GitLab\dExcelTools\dExcel\dExcel\dExcel.csproj'

print('Deleting unnecessary files...')
for file in os.listdir(release_build_path):
    if bool(re.match(r'(.+packed.+)|(.*\.pdb)|(dExcel-AddIn.xll)', file)):
        print(f'Deleting {file}')
        os.remove(os.path.join(release_build_path, file))

tree = ET.parse(dexcel_project_file_path)
version_number: str = tree.getroot().find('PropertyGroup').find('Version').text

print(f'dExcel version number: {version_number}')

print('Zipping files...')
with zipfile.ZipFile(rf'{version_number}.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
    zipdir(release_build_path, zipf)

source_path: str = os.path.join(os.getcwd(), f'{version_number}.zip')
target_path: str = os.path.join(r'\\ZAJNB010\Capital Markets 2\AQS Quants\dExcelTools\Releases', f'{version_number}.zip')
print(f'Copying {version_number}.zip to shared drive...')
shutil.move(source_path, target_path)
print('Process complete')