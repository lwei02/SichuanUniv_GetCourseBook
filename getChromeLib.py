# This script detects chrome version and downloads the corresponding chromedriver

import os, re
import sys
import requests
import zipfile
import keyboard

def getChromeVersion():
    pattern = r'\d+\.\d+\.\d+'
    cmd = r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version'
    stdout = os.popen(cmd).read()
    version = re.search(pattern, stdout)
    return version.group(0)

def downloadChromeDriver(version):
    mversion = version.split('.')[0]
    r = requests.get('https://chromedriver.storage.googleapis.com/') #xml
    pattern = r'{}\.\d+\.\d+\.\d+'.format(mversion)
    versions = re.findall(pattern, r.text)
    versions = sorted([*set(versions)], key=lambda x: list(map(int, x.split('.'))), reverse=True)
    version1 = versions[0]
    url = 'https://chromedriver.storage.googleapis.com/{}/chromedriver_win32.zip'.format(version1)
    r = requests.get(url)
    if not os.path.exists('./tmp'):
        os.makedirs('./tmp')
    os.chdir('./tmp')
    with open('chromedriver.zip', 'wb') as f:
        f.write(r.content)
    with zipfile.ZipFile('chromedriver.zip', 'r') as zip_ref:
        zip_ref.extractall()
    os.chdir('..')
    if not os.path.exists('./lib'):
        os.makedirs('./lib')
    if os.path.exists('./lib/chromedriver.exe'):
        cmd = r'lib\\chromedriver.exe --version'
        stdout = os.popen(cmd).read()
        v = stdout.split(' ')[1]
        print("\r\nThere's already a chromedriver, versioning ", end='')
        print(v, end='')
        print(', do you want to replace it? (Y/n)', end='')
        while True:
            if keyboard.is_pressed('y'):
                print('Y\r\nReplacing...', end='')
                break
            elif keyboard.is_pressed('n'):
                print('n\r\nExiting...', end='')
                sys.exit()
            elif keyboard.is_pressed('enter'):
                print('\r\n\r\nReplacing...', end='')
                break
        os.remove('./lib/chromedriver.exe')
    os.rename('./tmp/chromedriver.exe', './lib/chromedriver.exe')
    for i in os.listdir('./tmp'):
        os.remove('./tmp/{}'.format(i))
    os.rmdir('./tmp')


def main():
    print('\rGetting Chrome version...', end='')
    version = getChromeVersion()
    print('\r\nChrome version: {}'.format(version), end='')
    print('\r\nDownloading ChromeDriver...', end='')
    downloadChromeDriver(version)
    print('\r\nDone')

if __name__ == '__main__':
    main()