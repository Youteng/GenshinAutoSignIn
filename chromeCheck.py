import os
import json
import logging
import requests
import pathlib
from win32com import client as wincom_client

def read_json(file_path):
    with open(file_path, "r") as f:
        data = json.load(f)
    return data

def get_file_version(file_path):
    print('Checking Chrome version...')
    if not os.path.isfile(file_path):
        raise FileNotFoundError('File {0} is not found'.format(file_path))

    wincom_obj = wincom_client.Dispatch('Scripting.FileSystemObject')
    version = wincom_obj.GetFileVersion(file_path)
    print('The file version of {0} is {1}\n'.format(file_path, version))
    return version.strip().split('.')[0]

def write_json(file_path, data):
    with open(file_path, "w") as f:
        json.dump(data, f, indent=2)

def get_lastest_driver_version(chrome_version):
    url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{0}".format(chrome_version)
    response = requests.get(url)
    return response.text.strip()

def download_chrome_driver(version, dest_folder):
    url = "https://chromedriver.storage.googleapis.com/{0}/chromedriver_win32.zip".format(version)
    dest_path = os.path.join(dest_folder, os.path.basename(url))
    print("Downloading...")
    response = requests.get(url, stream=True, timeout=300)
    if response.status_code == 200:
        with open(dest_path, "wb") as f:
            f.write(response.content)
        print("Lastest chrome driver downloaded")
    else:
        raise Exception("Failed during downloading.\n")

current_path = pathlib.Path().resolve()
print("Current path: {0}\n".format(current_path))

configs = read_json('config.json')
file_path = configs['chrome_path']
driver_path = str(current_path) + '\\drivers\\'
chrome_driver_version = configs['chrome_driver_version']
current_chrome_version = get_file_version(file_path)


if(chrome_driver_version != current_chrome_version):
    print("Chrome version is {0}, and Chrome driver version is {1}".format(current_chrome_version, chrome_driver_version))
    print("Chrome driver needs to be updated")
    lastest_driver_version = get_lastest_driver_version(current_chrome_version)
    print("The lastest driver version is {0}".format(lastest_driver_version))
    download_chrome_driver(lastest_driver_version, driver_path)

print()