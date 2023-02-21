"""
This script is used to publish the latest release build of dExcel to the shared drive.
"""
import os
import re
import shutil
import socket
import xml.etree.ElementTree as ET
import zipfile

def ping(host):
    """
    Returns True if host responds to a ping request
    """
    import subprocess, platform

    # Ping parameters as function of OS
    ping_str = "-n 1" if  platform.system().lower()=="windows" else "-c 1"
    args = "ping " + " " + ping_str + " " + host
    need_sh = False if  platform.system().lower()=="windows" else True

    # Ping
    return subprocess.call(args, shell=need_sh) == 0

def zipdir(path, ziph):
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file),
                       os.path.relpath(os.path.join(root, file), path))


PING_HOST: str = 'https://gitlab.fsa-aks.deloitte.co.za'  

try:
    # The path on the local machine where the release build is created.
    release_build_path: str = r'C:\GitLab\dExcelTools\dExcel\dExcel\bin\Release\net6.0-windows'
    # The path on the shared drive containing the dExcel releases.
    shared_drive_releases_path: str = r'\\ZAJNB010\Capital Markets 2\AQS Quants\dExcelTools\Releases'
    dexcel_project_file_path: str = r'C:\GitLab\dExcelTools\dExcel\dExcel\dExcel.csproj'
    tree = ET.parse(dexcel_project_file_path)
    version_number: str = tree.getroot().find('PropertyGroup').find('Version').text

    # print(f'-----------------------------------------------------------')
    # print(f'Publishing dExcel version {version_number} to Shared Drive.')
    # print(f'-----------------------------------------------------------')
    # print(f'Checking connection to VPN...')

    # data = socket.gethostbyname(PING_HOST)

    print('Copy SQLite DLL for stats database...')
    sqlite_dll_source_path: str = os.path.join(release_build_path, r'runtimes\win-x64\native\SQLite.Interop.dll')
    sqlite_dll_target_path: str = os.path.join(release_build_path, "SQLite.Interop.dll")
    shutil.move(sqlite_dll_source_path, sqlite_dll_target_path)

    print('Deleting unnecessary, local files in release build folder...')

    for file in os.listdir(release_build_path):
        if bool(re.match(r'(.+packed.+)|(.*\.pdb)|(dExcel-AddIn.xll)', file)):
            print(f'\tDeleting {file}')
            os.remove(os.path.join(release_build_path, file))

    shutil.rmtree(os.path.join(release_build_path, 'runtimes'))  

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
            print(f'Overriding \'{version_number}.zip\' on shared drive...')
            shutil.move(source_path, target_path)
        else:
            print(f'Process aborted by user.')
    else:
        print(f'Copying \'{version_number}.zip\' to shared drive...')
        shutil.move(source_path, target_path)

    print('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
    print('                    Process Complete                        ')
    print('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
    print()

# except socket.error:
#     print('You are not connected to the VPN.')
#     print('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
#     print('         Failed to Publish dExcel to Shared Drive           ')
#     print('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')

except Exception as e:
    print(f'Unhandled exception {e} of type {type(e)} occurred.')
    print('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
    print('         Failed to Publish dExcel to Shared Drive           ')
    print('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')