import os
import subprocess
import sys
import time
from tkinter.filedialog import askopenfilename

from excel_reader import excel_data

rdp_options = """screen mode id:i:2
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

red = "\033[1;31m"
green = "\033[1;32m"
yellow = "\033[1;33m"
blue = "\033[1;34m"
magenta = "\033[1;35m"
cyan = "\033[1;36m"
white = "\033[1;37m"
reset = "\033[0m"


def showProgress(count, total, width=25, symbol='-', name=''):
    print("\r " + green + symbol * int(count / total * width) + red + symbol * (width - int(count / total * width)) +
          reset + f" {(count / total) * 100:.2f}% " + white + (f"[{name}]" if name else name) + reset, end="",
          flush=True)


def print_ascii_art(colour):
    print(colour + """
__________________   _____ _                _   _____       _    ___  ___      _             
| ___ \  _  \ ___ \ /  ___| |              | | /  __ \     | |   |  \/  |     | |            
| |_/ / | | | |_/ / \ `--.| |__   ___  _ __| |_| /  \/_   _| |_  | .  . | __ _| | _____ _ __ 
|    /| | | |  __/   `--. \ '_ \ / _ \| '__| __| |   | | | | __| | |\/| |/ _` | |/ / _ \ '__|
| |\ \| |/ /| |     /\__/ / | | | (_) | |  | |_| \__/\ |_| | |_  | |  | | (_| |   <  __/ |   
\_| \_|___/ \_|     \____/|_| |_|\___/|_|   \__|\____/\__,_|\__| \_|  |_/\__,_|_|\_\___|_|   
                                                                                                                                                                       
""" + reset)


if __name__ == '__main__':
    if sys.stdout.isatty():
        red = ""
        green = ""
        yellow = ""
        blue = ""
        magenta = ""
        cyan = ""
        white = ""
        reset = ""

    print_ascii_art(blue)
    company_path = askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])

    data = excel_data(company_path)
    company_path = os.path.split(company_path)[0]
    company_name = os.path.split(company_path)[1]

    print(green + "Saving to: " + reset + os.path.abspath(company_path), end="\n\n")

    start_time = time.time()
    count = 0
    for rdp_connection in data.rdp_connections:
        file_name = f"{company_name}-{rdp_connection.host}-" + str(rdp_connection.username) \
            .replace("\\", "_").replace("/", "_") + ".rdp"

        try:
            cmd = '("%s" | ConvertTo-SecureString -AsPlainText -Force) | ConvertFrom-SecureString;' % \
                  rdp_connection.password.replace('"', '""')
            hashed_pwd = subprocess.run(["powershell", "Invoke-Command -ScriptBlock", "{" + cmd + "}"],
                                        stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True).stdout

            showProgress(count=count, total=len(data.rdp_connections), name=file_name)

            rdp_file = open(company_path + "\\" + file_name, mode="w+", encoding="utf-8")
            rdp_file.write(rdp_options)
            rdp_file.write(f"full address:s:{rdp_connection.host}" + "\n")
            rdp_file.write(f"username:s:{rdp_connection.username}" + "\n")
            rdp_file.write(f"password 51:b:{hashed_pwd}" + "\n")
            rdp_file.close()
            count += 1
        except Exception as e:
            print(red, "|", str(e).replace("\n", ""), end=reset)

    showProgress(count=count, total=len(data.rdp_connections))

    ratio = count / len(data.rdp_connections)
    stat_color = green if ratio == 1 else yellow if ratio >= 0.5 else red
    stat_message = len(data.rdp_connections) if stat_color == green else f"{count}/{len(data.rdp_connections)}"

    time_passed = time.time() - start_time
    time_ratio = time_passed / len(data.rdp_connections) if count > 0 else 0
    time_colour = green if time_ratio < 0.1 else yellow if time_ratio < 0.2 else red

    print(f"\n\n{stat_color}{stat_message}{reset} RDP shortcut connections were created in "
          f"{time_colour}{time_passed:.2f}{reset} seconds")
