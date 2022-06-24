import os
import time

from excel_reader import excel_data

rdp_prams = """screen mode id:i:2
use multimon:i:0
desktopwidth:i:1920
desktopheight:i:1080
session bpp:i:32
winposstr:s:0,3,0,0,800,600
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
connection type:i:7
networkautodetect:i:1
bandwidthautodetect:i:1
displayconnectionbar:i:1
enableworkspacereconnect:i:0
disable wallpaper:i:0
allow font smoothing:i:0
allow desktop composition:i:0
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:0
disable cursor setting:i:0
bitmapcachepersistenable:i:1
audiomode:i:0
redirectprinters:i:1
redirectlocation:i:0
redirectcomports:i:0
redirectsmartcards:i:0
redirectclipboard:i:1
redirectposdevices:i:0
autoreconnection enabled:i:1
authentication level:i:2
prompt for credentials:i:0
negotiate security layer:i:1
remoteapplicationmode:i:0
alternate shell:s:
shell working directory:s:
gatewayhostname:s:
gatewayusagemethod:i:4
gatewaycredentialssource:i:4
gatewayprofileusagemethod:i:0
promptcredentialonce:i:0
gatewaybrokeringtype:i:0
use redirection server name:i:0
rdgiskdcproxy:i:0
kdcproxyname:s:
drivestoredirect:s:*
"""


def showProgress(count, total, width=25, symbol='-', name=''):
    print("\r   \033[1;32m" + symbol * int(count / total * width) + "\033[1;31m" + symbol * (width - int(
        count / total * width)) + f"\033[0m {(count / total) * 100:.2f}%\033[1;37m" + " " + name + "\033[0m", end="")


if __name__ == '__main__':
    print("\033[1;32m" + "Excel to RDP shortcut" + "\033[0m")

    company_path = input("\033[1;32m" + "Enter path to excel file or folder containing excel files" + "\033[0m" + ":")
    if os.path.isdir(company_path):
        excels = list(filter(lambda f: ".xlsx" in str(f), os.listdir(company_path)))
        if len(excels) == 0:
            print("Error - Gave incorrect path to excel file")
            exit(-1)
        company_path += "\\" + excels[0]

    data = excel_data(company_path)
    company_path = os.path.split(company_path)[0]
    company_name = os.path.split(company_path)[1]

    start_time = time.time()
    count = 0
    for rdp_connection in data.rdp_connections:
        file_name = f"{company_name}-{rdp_connection.host}-" + str(rdp_connection.username) \
            .replace("\\", "_").replace("/", "_") + ".rdp"

        hashed_pwd = os.popen(f"cryptRDP5.exe {rdp_connection.password}").read()

        showProgress(count=count, total=len(data.rdp_connections), name=file_name)

        rdp_file = open(company_path + "\\" + file_name, mode="w+", encoding="utf-8")
        rdp_file.write(rdp_prams)
        rdp_file.write(f"full address:s:{rdp_connection.host}" + "\n")
        rdp_file.write(f"username:s:{rdp_connection.username}" + "\n")
        rdp_file.write(f"password 51:b:{hashed_pwd}" + "\n")
        rdp_file.close()
        count += 1

    showProgress(count=count, total=len(data.rdp_connections))

    print(f"\n\n\033[1;32m{len(data.rdp_connections)} RDP shortcut connections were created in {time.time() - start_time:.2f} seconds\033[0m")
