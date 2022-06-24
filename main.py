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

if __name__ == '__main__':
    company_path = input("Enter Path of wanted Excel file: ")
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
    for rdp_connection in data.rdp_connections:
        print(rdp_connection)
        print()

        file_name = f"{company_name}-{rdp_connection.host}-" + str(rdp_connection.username) \
            .replace("\\", "_").replace("/", "_") + ".rdp"

        hashed_pwd = os.popen(f"cryptRDP5.exe {rdp_connection.password}").read()

        rdp_file = open(company_path + "\\" + file_name, mode="w+", encoding="utf-8")
        rdp_file.write(rdp_prams)
        rdp_file.write(f"full address:s:{rdp_connection.host}" + "\n")
        rdp_file.write(f"username:s:{rdp_connection.username}" + "\n")
        rdp_file.write(f"password 51:b:{hashed_pwd}" + "\n")
        rdp_file.close()

    print(f"Finished... Done {len(data.rdp_connections)} Rdp Shortcuts in {time.time() - start_time:.2f} s")
