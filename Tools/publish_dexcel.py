"""
This script is used to publish the latest release build of dExcel to the shared drive.
"""
from colorama import init as colorama_init
from colorama import Fore
from colorama import Style
import os
import re
import shutil
import socket
import xml.etree.ElementTree as ET
import zipfile

# This is the URL to the FSA instance of gitlab. It is used by the function 'ping'.
# The idea is that if the DNS cannot resolve the IP of this host then the user is not 
# connected to the VPN.
PING_HOST: str = 'gitlab.fsa-aks.deloitte.co.za'  

def print_process(message: str) -> None:
    """
    Prints a message indicating that a process (such as copying a file, deleting a file etc.) is being kicked off
    in a standardised forward.

    :param message: The message to print.
    :return: None
    """
    print(f' - {message}... ', end='')

def print_warning(message: str) -> None:
    """
    Prints a warning message in a standardised forward.

    :param message: The message to print.
    :return: None
    """
    print(f'{Fore.YELLOW}{Style.BRIGHT}', end='')
    print(f' • {message}\n')
    print(f'{Style.RESET_ALL}', end='')

def ok_message() -> None:
    """
    Prints a stylized 'OK'. Used after a process has succeeded.

    :return: None
    """
    print(f'[{Fore.LIGHTGREEN_EX}{Style.BRIGHT} OK {Style.RESET_ALL}]\n') 

def error_message() -> None:
    """
    Prints a stylized 'ERROR'. Used after a process has failed.

    :return: None
    """
    print(f'[{Fore.LIGHTRED_EX}{Style.BRIGHT} ERROR {Style.RESET_ALL}]\n')

def failed_to_publish_message() -> None:
    """
    Prints a message that ∂Excel failed to publish to the shared drive.
    This is used if the program is been fatally stopped.
    
    :return: None
    """
    print(f'{Fore.LIGHTRED_EX}{Style.BRIGHT}')
    print(f'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
    print(f'         Failed to Publish ∂Excel to Shared Drive           ')
    print(f'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
    print(f'{Style.RESET_ALL}')

def successfully_published_dexcel_message() -> None:
    """
    Prints a message that ∂Excel successfully to publish to the shared drive.
    This is used once all processes have successfully completed.

    :return: None
    """
    print(f'{Fore.GREEN}{Style.BRIGHT}')
    print(f'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
    print(f'       Successfully Published ∂Excel to Shared Drive        ')
    print(f'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
    print(f'{Style.RESET_ALL}')

def ping(host: str) -> bool:
    """
    Returns True if it can connect to the host else False.

    :param host: The URL string for the host.
    :return: True if it can ping the host, otherwise False.
    """
    try:
        print_process('Checking connection to VPN')
        socket.gethostbyname(host)
        ok_message()
        return True

    except socket.error as e:
        error_message()
        print(f' - Socket error encountered: {Fore.LIGHTRED_EX}{Style.BRIGHT}{e}{Style.RESET_ALL}')
        print_warning(f'Are you connected to the VPN?')
        failed_to_publish_message()
        return False

def zip_directory(path, ziph) -> None:
    """
    Zips/compresses a directory.

    :param path: The path to the directory.
    :param ziph: The location of the zip file.
    :return: None 
    """
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file),
                    os.path.relpath(os.path.join(root, file), path))

try:
    colorama_init(convert=True)
    # The path on the local machine where the release build is created.
    release_build_path: str = r'C:\GitLab\dExcelTools\dExcel\dExcel\bin\Release\net6.0-windows'
    # The path on the shared drive containing the dExcel releases.
    shared_drive_releases_path: str = r'\\ZAJNB010\Capital Markets 2\AQS Quants\dExcelTools\Releases'
    dexcel_project_file_path: str = r'C:\GitLab\dExcelTools\dExcel\dExcel\dExcel.csproj'
    tree = ET.parse(dexcel_project_file_path)

    version_number: str = tree.getroot().find('PropertyGroup').find('Version').text

    zip_file: str = f'{Fore.GREEN}{Style.BRIGHT}\'{version_number}.zip\'{Style.RESET_ALL}'

    print(f'{Fore.LIGHTGREEN_EX}{Style.BRIGHT}')
    print(f'-----------------------------------------------------------')
    print(f'             Publishing ∂Excel to Shared Drive             ')
    print(f'-----------------------------------------------------------')
    print(f'{Style.RESET_ALL}')

    print(f' • ∂Excel version number: {Fore.GREEN}{Style.BRIGHT}{version_number}{Style.RESET_ALL}\n')

    if ping(PING_HOST):
        print_process('Copying SQLite DLL for stats database')
        sqlite_dll_source_path: str = os.path.join(release_build_path, r'runtimes\win-x64\native\SQLite.Interop.dll')
        sqlite_dll_target_path: str = os.path.join(release_build_path, 'SQLite.Interop.dll')
        ok_message()

        if not os.path.exists(sqlite_dll_target_path):
            shutil.copy(sqlite_dll_source_path, sqlite_dll_target_path)

        print_process('Deleting unnecessary, local files in release, build folder')
        deleted_files_count: int = 0
        for file in os.listdir(release_build_path):
            if bool(re.match(r'(.+packed.+)|(.*\.pdb)|(dExcel-AddIn.xll)', file)):
                deleted_files_count += 1
                print(f'\tDeleting {file} ', end='')
                os.remove(os.path.join(release_build_path, file))
                ok_message()

        if deleted_files_count == 0:
            ok_message()

        sqllite_runtimes_path: str = os.path.join(release_build_path, 'runtimes')
        if os.path.exists(sqllite_runtimes_path):
            print_process('Deleting unnecessary, SQLite runtimes folder')
            shutil.rmtree(sqllite_runtimes_path)  
            ok_message()

        print_process('Zipping local files')
        with zipfile.ZipFile(rf'{version_number}.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
            zip_directory(release_build_path, zipf)

        ok_message()

        source_path: str = os.path.join(os.getcwd(), f'{version_number}.zip')
        target_path: str = os.path.join(shared_drive_releases_path, f'{version_number}.zip')
        file_stats: os.stat_result = os.stat(source_path)
        file_size: str = f'({file_stats.st_size / (1024 * 1024):.2f}MB)'

        if os.path.exists(target_path):
            print_warning(f'File \'{target_path}\' already exists!')
            yes_no: str = f'[{Fore.GREEN}{Style.BRIGHT}y{Fore.LIGHTBLACK_EX}/{Fore.YELLOW}n{Fore.LIGHTBLACK_EX}]{Style.RESET_ALL}'
            print(f'{Fore.LIGHTBLACK_EX}{Style.BRIGHT}', end='')
            option_to_overwrite_file: str = input(f'   □ Would you like to overwrite the file? {yes_no} ')
            print(f'{Style.RESET_ALL}', end='')
            
            if option_to_overwrite_file.upper() == 'Y':
                print()
                print_process(f'Overriding {zip_file} {file_size} on shared drive')
                shutil.copy(source_path, target_path)
                ok_message()
                successfully_published_dexcel_message()
            else:
                print(f'{Fore.YELLOW}{Style.BRIGHT}')
                print(f'-----------------------------------------------------------')
                print(f'                  Process Aborted by User                  ')
                print(f'-----------------------------------------------------------')
                print()
                print(f'{Style.RESET_ALL}', end='')
        else:
            print_process(f'Copying {zip_file} {file_size} to shared drive')
            shutil.copy(source_path, target_path)
            ok_message()

except Exception as e:
    error_message()
    print(f'Unhandled exception {e} of type {type(e)} occurred.')
    failed_to_publish_message()
