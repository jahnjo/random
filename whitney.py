# Author: John Foster (@juf317)

# config variables
# username = juf317


import subprocess
import importlib.util
import sys
import os
import win32com.shell.shell as shell

import getpass
username = getpass.getuser() # getpass.getuser() should get your username and put it in lowercase letters. If this didn't work, comment THIS username line out and uncomment the ABOVE username line and manually set your username there.

import socket
hostname = socket.gethostname().lower() # socket.gethostname().lower() should get your hostname and put it in lowercase letters. If this didn't work, comment THIS hostname line out and uncomment the ABOVE hostname line and manually set your hostname there.

selenium_directory = 'selenium'

make_selenium_dir = 'mkdir selenium'

mozilla_geckodriver_download_url = 'https://github.com/mozilla/geckodriver/releases/download/v0.17.0/geckodriver-v0.17.0-win64.zip' 

IE_server_download_url = 'http://selenium-release.storage.googleapis.com/3.4/IEDriverServer_x64_3.4.0.zip' 

se_standalone_server_download = 'http://selenium-release.storage.googleapis.com/3.4/selenium-server-standalone-3.4.0.jar'

json_config_download = 'https://confluence.sl-cloud.arl.psu.edu/download/attachments/38142119/config.json?version=1&modificationDate=1497972348118&api=v2'

wget_install_command = 'pip install wget'

remove_selenium_command = 'rmdir /s /q selenium'

start_server = 'start java -jar selenium-server.jar -role hub'

start_node = 'start java -Dwebdriver.gecko.driver=./drivers/geckodriver.exe -Dwebdriver.gecko.driver=./drivers/IEdriverserver.exe -jar selenium-server.jar -role node -nodeConfig config.json'.format(username)

dipndots = ["setuptools", "appdirs", "autopep8", "cffi", "colorama", "cryptography", "idna", "packaging", "paramiko", "pep8", "pyasn1", "pycparser", "pycparser", "pyparsing", "six", "termcolor", "wget"] 

pip_path = sys.exec_prefix + '\Scripts\pip'

python_path = sys.executable


def download_pip():
    subprocess.call('python get-pip.py', shell=True)

if(subprocess.call('pip --version', shell=True)):
    print('\nIt appears you dont have pip installed, I will install it for you\n')
    download_pip()
    
    subprocess.call('{} whitney.py'.format(python_path), shell=True)
    quit()



def dependencies():

    print('\n\nINSTALLING DEPENDENCIES OR CONFIRMING DEPENDENCIES ARE ALREADY INSTALLED\n\n')

    new_dips = 0;
    for dip in dipndots:

        if not importlib.util.find_spec(dip):
            new_dips += 1
            subprocess.call('{} install -U {}'.format(pip_path, dip),shell=True)

    return new_dips <= 1 

if (not dependencies()):
    subprocess.call('{} install pypiwin32'.format(pip_path), shell=True)
    subprocess.call('{} whitney.py'.format(python_path), shell=True)
    quit()


import urllib.request
import zipfile
from pathlib import Path
import shutil
import win32com.shell.shell as shell
import xml.etree.ElementTree as ET
import webbrowser
from shutil import move
from shutil import copy
import colorama
from win32com.client import Dispatch
from termcolor import *
import wget
from shutil import copyfile

temp_dir = os.getcwd()

# global functions

def wait_for_user_input(required_message='Yes'):
    '''Prompts user for input'''
    while True:
        print('\nType "{}" when you are ready to continue.'.format(required_message))
        print('Type CTRL+C to quit.')
        user_input = input('> ')
        if (user_input == required_message):
            break

def wait_for_user_input_options(required_message_1='Yes', required_message_2='No',required_message_3='Maybe'):
    '''Prompts user for input'''
    while True:
        print('\nType "{}", "{}", or "{}" when you are ready to continue.'.format(required_message_1,required_message_2,required_message_3))
        print('Type CTRL+C to quit.')
        user_input = input('> ')
        if (user_input == required_message_1):
            wait_for_user_input_options.option = required_message_1
            break
        elif (user_input == required_message_2):
            wait_for_user_input_options.option = required_message_2
            break
        elif (user_input == required_message_3):
            wait_for_user_input_options.option = required_message_3
            break

def wait_for_user_input_two_options(required_message_1='Yes', required_message_2='No'):
    '''Prompts user for input'''
    while True:
        print('\nType "{}", or "{}" when you are ready to continue.'.format(required_message_1,required_message_2))
        print('Type CTRL+C to quit.')
        user_input = input('> ')
        if (user_input == required_message_1):
            wait_for_user_input_options.option = required_message_1
            break
        elif (user_input == required_message_2):
            wait_for_user_input_options.option = required_message_2
            break



# setup functions

def init():


    if os.path.exists('./selenium'):
        
        print('\n\nEverything has already been downloaded and installed. Type "s" to start the server or if you want to reinstall type "r".')
        wait_for_user_input_two_options('s', 'r')
        if (wait_for_user_input_options.option == 's'):
            os.chdir(selenium_directory)
            subprocess.call(start_server,shell=True)
            subprocess.call(start_node,shell=True)
        elif(wait_for_user_input_options.option == 'r'):
            cprint('\n\nREINSTALLING SELENIUM SERVER', 'cyan')
            dependencies()
            work_horse()
    
    else:
        '''Start of the setup and ask the user to choose which browser they will use to test'''
        colorama.init()
        
        print('\n\n\nWelcome to whitney.py, where I will help you setup a selenium server, before you start anything open up README and make sure python is downloaded\nProceed (y/n)?')
        wait_for_user_input_two_options('y', 'n')
        if(wait_for_user_input_options.option == 'y'):
            work_horse()
        elif (wait_for_user_input_options.option == 'n'):
            quit()

def work_horse():

    subprocess.call(remove_selenium_command, shell=True) # removes any existing selenium files in home dir 

    subprocess.call('mkdir selenium', shell=True) # makes a new selenium dir
        
    move_json_config() # moves json config file to selenium dir

    os.chdir(selenium_directory) # change into selenium dir

    print('\nSelenium Testing Setup \n')
    print('\nWhich browser(s) would you like to run tests for?\n1: Both\n2: Firefox\n3: Internet Explorer\n(1,2,3)?')
    wait_for_user_input_options('1', '2', '3')

    if (wait_for_user_input_options.option == '1'):
        install_geckodriver()
        unzip_file('./geckodriver.zip')
        install_IE_driver()
        unzip_file('./IEDriverServer.zip')
    elif (wait_for_user_input_options.option == '2'):
        install_geckodriver()
        unzip_file('./geckodriver.zip')
    elif (wait_for_user_input_options.option == '3'):
        install_IE_driver()
        unzip_file('./IEDriverServer.zip')
       
    install_se_server()
    subprocess.call(start_server,shell=True)
    subprocess.call(start_node,shell=True)
    


def install_geckodriver():
    '''Opens url to geckodriver download page and lets user download and install geckodriver'''
    cprint('\n\nINSTALL GECKODRIVER', 'cyan')
    print('\n\nIn order to run tests on Firefox we have to download geckodriver')
    print('I will download Geckodriver for you')

    wget.download(mozilla_geckodriver_download_url, out='./geckodriver.zip')

def install_IE_driver():
    '''Opens url to ie driver download page and lets user download and install ie driver'''
    cprint('\n\nINSTALL INTERNET EXPLORER DRIVER', 'cyan')
    print('\n\nIn order to run tests on Internet Explorer we have to download the Internet Explorer Driver Server')
    print('I will download the Internet Explorer Driver Server for you')

    wget.download(IE_server_download_url, out='./IEDriverServer.zip')

def install_se_server():
    '''Opens url to selenium standalone server and lets user download and install it'''
    cprint('\n\nINSTALL SELENIUM STANDALONE SERVER', 'cyan')
    print('\n\nWere going to download the selenium standalone server jar')
    print('I will download the Selenium Standalone Server for you')

    wget.download(se_standalone_server_download, out='./selenium-server.jar')

def move_json_config():
    os.chdir(temp_dir)
    copyfile("config.json","./selenium/config.json".format(username))

def unzip_file(zip_file_path):
    '''Extracts a Zip file to a folder named after the Zip file in the same directory'''
    zip_ref = zipfile.ZipFile(zip_file_path, 'r')
    zip_ref.extractall('drivers')
    zip_ref.close()

print('''\n\n

MMMMMMMMMMMNNNNdyoo/:----.-``.....```-:/-````.-:/:...`-os++/-.`-/+//:ss+:.+hmdd
MMMMMMMMMmmNmdhsoooo+/..--:-`-``..``  .-+/`.--oyy/:-```+syo/:-://----/so:-:osoo
MMMMMMMMmMMdmNhdssso:--.`-`//:-::``.``.:+o::-/hmhs+/-`.::/--o/--:::::::/:.`/osh
MMMMMMNNmMNdh//oso//o/:..``-o/os-..-..:/oso+-+shmyso+../-------.-/:++.-:--::os+
MMMMMMddmNmhs-:///++ss:.+-:`+oho:-...-:///ssssssddyo+.`-//+:/o+.-:/:.`./:..-:/o
MMMMMMmmmmdhs//+//::/:/-//s:--s+/:.-.-----:ohdyodmds+-.`-:::--:/o+:`.`.:....`/h
MMMMMMNmNNdh+s+-+://::/:/o/-.`:-:....``..--+hdhsdNNNd+-:os:.-//++/:.`.-:/::--:h
MMMMMMMNNsoo/sy--/+//:---..```.`::-::..-`-:+hdyydNNNo:-:+/./+oososo//`.://+oo//
MMMMMMmmmo++//oo/..--...-:..```-ss+s+//:-++://+yhdds-.:ss//osso++oosy/``.----..
MMMNNNhhhdyyo+//::....-:::-....-syo+::::::/++syyyys+-.+ddysyyyyyhdmmmd-   ``-::
MNNNmmmhhyhys++//-..``.`````--..:::-.-:--/osyyyss+/::/sddhyyyyyyhyyydms`  `.-/+
MNmmmmddhysoo+/::.`.---:+:-..-.......-.-:+sshyysoo/::odNmhyhhsssooo++oy````.---
MMNmmmdhyss+:---.``-.:+///:--..`````...--:--/://+oooosmNNddhs-.-:-/.-:o.``.-::-
MNMNmmyosysso/-----....-....```````.---..-:/oyhsosyhhhmNNNmmhyyyhhdhyyd+```..--
MNNMNm+:::+o+//:/:///:-::-:``````.:/+ossyhhhdmmmmmddhdmNNNNmmNNmNmmmmmmm:...::/
MMNmhs/::://////+//++ooso/.````.-/osyddmmmmmNNNNmmmdddmNNNNNNNNNNNNNNNNNm/---:+
MMMNmhso+-/+/:://++++o++:.`````-:/oyhhdmmNNNNNNNNmdddddmNMNNmmNNNNNNNNNNNd//:os
MMMMNNmds:/+/:--..--...-:.`````.-:+osyhdmNNNNNNNmoss+/+oyhhy++ymmmNNNNNNNm-/shy
MMMMMNNh+:++//:-..``.`.-..`````..-:+osyhdmmNNNmmms//::--:/+osyhmmmmmmmmmmm-+o+s
MMMMMNNmdhho/:--..-`::::-.``````..-:/+osyhdmmmmdddddddmmddmmmmmmmmmmmmddmd-osoh
MMMMMMNmmmmdhsso:/o..:+y/.` `````.---:/+osyhhhyyyyyhhhddmmddhhyhdmmmddhdms/ossy
MMMMMMMMNmNdyyhy++s/.:/oo:..-````..---::/+oyo+++///+++sshhshhhso/oyhdddmm/:/ohm
MMMMMMMMMMMNmmmmhsyo:--.....-.```.--::::/+shy/-..-osdhNmNMNMMNNms/ohdmmNd::/smN
MMMMMMMMMMMMMNddmdhs/:-:-.....```.-:::::/+oyddhsosyyyydmNNNNNmdhhdmmmmmNhsyhNMM
MMMMMMMMMMMMMNyhyhdy+//:-..```  ``-:///::/+shdhhhysssosyhhhhhhhhdmmmmmNNssshNMM
MMMMMMMMMMMMMNhdhsho+/:---.```   `.:////://+syyyhhhhhhddmmmmmmmmmmmmmmNm+ymMMMM
MMMMMMMMMMMMMNmmmhh+-.`````````   `-/////://+ossyyhddmmNNNNNNNNNmmmmmNMNNNMMMMM
MMMMMMMMMMMMMMMNmmmy/-..```..-`   `.-/+o+/:://+ossyhdmmNNNNNNNNNmmmmmNMMMMMMMMM
MMMMMMMMMMMMMMMMNMMNh+:-...--.``  `.-:/+++/::::/++osyhdmNNNNNmmmdddNNNMMMMMMMMM
MMMMMMMMMMMMMMMMMMMMMdyso:---.``  `.:::::--------:://+syyyyyyyyssydNMMMMMMMMMMM
MMMMMMMMMMMMMMMMMMMMMMdyyhs+oo.`  `//:----.............--------:sdMMMMMMMMMMMMM
MMMMMMMMMMMMMMMMMMMMMMMMNNmmhys:` -my+:-.......`````````````..:hmNMMMMMMMMMMMMM
MMMMMMMMMMMMMMMMMMMMMMMMMMMMNmy/.``Nmy/--...````````````````..:dhmNNNNmdmNMMMMM
MMMMMMMMMMMMMMMMMMMMMMMMMMMMMNmhso/dNh+:---..`````````````..-:/syh:::::+oshdmNM
MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMd:sNh+//::-..``````````..--:/+syy..-:::/+oyyyd
MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMNysy``syo////::--.........--:/++osyh:..-::://+oso
MMMMMMMMMMMMMMMMMMMMMMMMMMMMMh-`-o. `-//+++///:::::::--::/+oossyhh/...--:////+o
MMMMMMMMMMMMMMMMMMMMMMMMMMMNh+:.``.````.:++++++//++++///+oossyyhdhs.`.---:+////




''')



init()
quit()
