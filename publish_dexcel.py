# This script is used to publish the latest release build of dExcel to the shared drive.
import os
import re
import shutil
import xml.etree.ElementTree as ET
import zipfile


def zipdir(path, ziph):
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file),
                       os.path.relpath(os.path.join(root, file), path))

# The path on the local machine where the release build is created.
release_build_path: str = r'C:\GitLab\dExcelTools\dExcel\dExcel\bin\Release\net6.0-windows'
# The path on the shared drive containing the dExcel releases.
shared_drive_releases_path: str = r'\\ZAJNB010\Capital Markets 2\AQS Quants\dExcelTools\Releases'
dexcel_project_file_path: str = r'C:\GitLab\dExcelTools\dExcel\dExcel\dExcel.csproj'

tree = ET.parse(dexcel_project_file_path)
version_number: str = tree.getroot().find('PropertyGroup').find('Version').text

print('-----------------------------------------------------------')
print('Publishing dExcel version {version_number} to Shared Drive.')
print('-----------------------------------------------------------')
print('Deleting unnecessary, local files in release build folder...')

for file in os.listdir(release_build_path):
    if bool(re.match(r'(.+packed.+)|(.*\.pdb)|(dExcel-AddIn.xll)', file)):
        print(f'\tDeleting {file}')
        os.remove(os.path.join(release_build_path, file))


print(f'dExcel version number: {version_number}')

print('Zipping local files...')
with zipfile.ZipFile(rf'{version_number}.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
    zipdir(release_build_path, zipf)

source_path: str = os.path.join(os.getcwd(), f'{version_number}.zip')
target_path: str = os.path.join(shared_drive_releases_path, f'{version_number}.zip')

if os.path.exists(target_path):
    print(f'File \'{target_path}\' already exists.')
    option_to_overwrite_file: str = input(f'Would you like to overwrite the file? \'y/n\' ')

    if option_to_overwrite_file.upper() == 'Y':
        print(f'Copying \'{version_number}.zip\' to shared drive...')
        shutil.move(source_path, target_path)
    else:
        print(f'Process aborted by user.') 
        
print('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
print('Process Complete')
print('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
print()
