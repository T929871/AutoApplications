import win32con
import winreg
import os
# import shutil
import unicodedata
# import time
import PySimpleGUI as sg
import getpass
import sys
#import getopt

# import threading

import urllib.request
import shutil
# import difflib
import re
from fuzzywuzzy import fuzz
# from fuzzywuzzy import process

# pyinstaller -wF -F --uac-admin --uac-uiaccess AutoApplications.py


def getAsset():
    layout = [[sg.T('Enter old Asset #')],
              [sg.InputText(), sg.Checkbox('Append ".corp.ads"', default=False, key='corp')],
              [sg.Submit(size=(18, 3), font='Arial 12', button_color=('black', 'orange')), sg.Cancel(size=(18, 3), font='Arial 12')]]

    window = sg.Window('Auto Applications Tool', layout)

    event, text_input = window.Read()
    window.Close()

    if event is None or event == 'Exit' or event == 'Cancel':
        sys.exit()
    asset = text_input[0]
    if text_input['corp'] is True:
        asset = append_corpads(asset)
    if not text_input:
        sys.exit()
    #print(asset)
    return asset

# appends the full_list with any software found
def getkeys(newAsset, full_list, hive, flag):
    aReg = winreg.ConnectRegistry(newAsset, hive)
    aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", 0, win32con.KEY_READ | flag)

    count_subkey = winreg.QueryInfoKey(aKey)[0]
    for i in range(count_subkey):
        try:
            asubkey_name = winreg.EnumKey(aKey, i)
            asubkey = winreg.OpenKey(aKey, asubkey_name)
            val = winreg.QueryValueEx(asubkey, "DisplayName")[0]
            val = unicodedata.normalize('NFKD', val).encode('ascii', 'ignore')
            val = val.decode('ascii')

            # preliminary filtering
            ignoreList = ['runtime', 'redistributable', 'installer', 'setting', 'updater', 'update for ', 'extension',
                          'proofing', 'flash player', 'add-in', 'driver', 'activex', 'anyconnect', 'microsoft access',
                          'microsoft dcf', 'microsoft excel', 'microsoft groove', 'infopath', 'lync', 'components',
                          'office osm', 'office shared', 'onenote', 'outlook', 'powerpoint', 'microsoft word', 'parser',
                          'myid', 'publisher', 'silverlight', 'framework', 'policy', 'plug-in', 'help', 'webcam',
                          'driveguard', 'service pack', 'support', 'microsoft office', 'hp ', 'intel ', 'intel(r) ',
                          'apple ', 'viewer', 'hotfix']  # lowercase
            skip = False
            for item in ignoreList:
                if item in val.lower():
                    skip = True
            if skip == False:
                full_list.append(val)
            else:
                continue
                # skip

        except EnvironmentError:
            continue

    full_list.sort()  # alpha sort

    return full_list


def getApps(full_list):
    col = [[sg.Submit(size=(18, 3), font='Arial 12', button_color=('black', 'orange'))],
           [sg.Cancel(size=(18, 3), font='Arial 12')]]

    layout = [
        [sg.Listbox(full_list, font=("Helvetica", 14), pad=(3, 2), select_mode='multiple', size=(45, 25)),
         sg.Column(col)]
        # [sg.Button("Select All")], [sg.Button("Clear Selection")], #not really that useful

    ]

    window = sg.Window('Select Applications to Install', layout, default_element_size=(40, 1),
                       grab_anywhere=False).Finalize()
    # window.Maximize()

    while True:
        event, list = window.Read()
        if event is None or event == 'Exit' or event == 'Cancel':
            sys.exit()

        # if none selected
        elif event == 'Submit' and not list[0]:

            layout = [[sg.T('No Applications Selected!')], [sg.Cancel()]]
            pop = sg.Window(':O', layout)
            pop.Read()
            pop.Close()

        else:
            # print(list[0])
            window.Close()
            return list[0]



#
def fileWrite(list, file):
    for item in list:
        file.write(item + '\r\n')
    return


# search software from library
def sw_search(software):
    # get first 4 letters
    sw4 = software[:4].lower()
    # remove spaces
    sw4 = sw4.replace(" ", "")
    # special cases *********************************************************************************************************
    if sw4 == 'micr':
        sw4 = 'msft'

    urlstr = "http://142.174.118.47/swlibrary/dbsearch.asp?search=" + sw4

    # print('Search token \"' + sw4 + '"\n')

    parsePage(urlstr, software, sw4)


# turn page into text for processing
def parsePage(url, software, sw4):
    pagehtml = open("pagehtml.txt", "wb")

    page = urllib.request.urlopen(url)
    shutil.copyfileobj(page, pagehtml)

    with open('pagehtml.txt', encoding='cp1252') as f:
        data = f.read()
    with open('pagehtml.txt', 'w', encoding='utf8') as f:
        f.write(data)
    pagehtml.close()

    matchindex = 0
    res = readPage(software, sw4, 0)
    while True:
        if res is 0:
            break
        elif res is 1:
            matchindex = matchindex + 1
        elif res is 2:
            matchindex = 0
        res = readPage(software, sw4, matchindex)
    return


def readPage(software, sw4, matchindex):
    with open('pagehtml.txt', encoding='utf8', errors='ignore') as f:
        pagehtml = f.read()
        # print(pagehtml)

        # fine tuning here
        # ************************************************************************************************************************
        # SPECIAL CASES
        # ---bluezone
        if "BlueZone " in software:
            software = software.replace('BlueZone ', 'BlueZone Desktop')

        # remove -
        software = software.replace('-', '')
        # replace spaces with -
        software = software.replace(' ', '-')
        # remove brackets
        software = software.replace('(', '')
        software = software.replace(')', '')
        # prefer APP-packages
        software = "APP-" + software
        software = software.rstrip()

        # path directly to directory
        path = '		<td width="348"><font face="Arial" size="2"><a href="file://\\\\corp.ads\\software\\release\\apps\\' + software + '"> \\\\corp.ads\\software\\release\\apps\\' + software + '</a></font></td>'
        # print(path)
        pathCmp = re.sub('\W+', '', path)
        # ************************************************************************************************************************
        page_lines = pagehtml.splitlines()

        score = []
        #cmper = 'null'
        for line in page_lines:
            cmper = re.sub('\W+', '', line)
            score.append(fuzz.ratio(path, cmper))
            #print(line)

        # have to do this or same score will return only first index
        visited = []
        for i in range(len(score)):
            iter(i, score, visited)
        #print(score)


        if matchindex is 0:
            index = score.index(max(score))
        else:
            scoreset = frozenset(score)
            scoreset = sorted(score, reverse=True)
            index = score.index(scoreset[matchindex])






        html = page_lines[index]
        #print(html)

        # print()
        try:
            url = re.findall(r'file\:\/\/(.*?)"\>', html)
            path = url[0]
        except IndexError:
            layout = [[sg.T("Looking for: \n" + software[4:], font=("Helvetica", 20))],
                      [sg.T('\nSearch token: \"' + sw4 + '"\n')],
                      [sg.T("*** No Results ***\n\n")],
                      [sg.OK(button_color=('black', 'orange'), size=(18, 3), font='Arial 12'),
                       sg.RealtimeButton('Open Software Library', size=(18, 3), font='Arial 12')]
                      ]

            window = sg.Window(software, layout, auto_size_text=True)

            while (True):
                # This is the code that reads and updates your window
                event, values = window.Read()
                if event is "Open Software Library":
                    try:
                        os.startfile("http://142.174.118.47/swlibrary/dbsearch.asp?search=" + sw4)
                    except:
                        pass
                if event == 'Quit' or values is None:
                    sys.exit()
                if event is "OK":
                    window.close()
                    return 0

            return
        if "rosoftApps" in path[33:]:
            # print("*** Not in Software Library ***\n\n")
            layout = [[sg.T("Looking for: \n" + software[4:], font=("Helvetica", 20))],
                      [sg.T('\nSearch token: \"' + sw4 + '"\n'), sg.Btn("Best match", font='Arial 10')],
                      [sg.T("*** No results ***\n\n")],
                      [sg.OK(button_color=('black', 'orange'), size=(18, 3), font='Arial 12'),
                       sg.RealtimeButton('Open Software Library', size=(18, 3), font='Arial 12')]
                      ]

            window = sg.Window(software, layout, auto_size_text=True)

            while (True):
                event, values = window.Read()

                if event is "Best match":
                    window.close()
                    return 2

                if event is "Open Software Library":
                    try:
                        os.startfile("http://142.174.118.47/swlibrary/dbsearch.asp?search=" + sw4)
                    except:
                        pass
                if event == 'Quit' or values is None:
                    sys.exit()
                if event is "OK":
                    window.close()
                    return 0

            return

        layout = [[sg.T("Looking for: \n" + software[4:], font=("Helvetica", 20))],
                  [sg.T('\nSearch token: \"' + sw4 + '"\n'), sg.Btn("Best match", font='Arial 10'), sg.Btn("Next match", font='Arial 10')],
                  [sg.T("Found: \n" + path[33:] + "\n", font=("Helvetica", 20))],
                  [sg.RealtimeButton('Install', size=(18, 3), font='Arial 12'), sg.RealtimeButton('Uninstall', size=(18, 3), font='Arial 12')],
                  [sg.RealtimeButton('Next', size=(18, 3), font='Arial 12', button_color=('black', 'orange')),
                   sg.RealtimeButton('Open Software Library', size=(18, 3), font='Arial 12')]
                  ]

        window = sg.Window(software, layout, auto_size_text=True)

        while (True):
            # This is the code that reads and updates your window
            event, values = window.Read()
            #print(event)
            if event is "Best match":
                window.close()
                return(2)
            if event is "Next match":
                window.close()
                return(1)
            if event is "Install":
                # look for install script
                script = '-install.cmd'
                exe = '_install.exe'
                # print(path)
                listdir = os.listdir(path)
                fileList = []
                found = False

                for item in range(len(listdir)):
                    fileList.append(listdir.pop().lower())

                for file in fileList:
                    # print(file)
                    if script in file:
                        # print("\nInstall Script found.")
                        # run script directly
                        try:
                            os.startfile(path + '\\' + file)
                        except:
                            pass

                        found = True
                        break
                    elif exe in file:
                        # print("\nInstall Executable found.")
                        # run script directly
                        try:
                            os.startfile(path + '\\' + file)
                        except:
                            pass
                        # print("\n")
                        found = True
                        break
                if found == False:
                    # open directory
                    # print("Manual Install Required.\n")
                    try:
                        os.startfile(path)
                    except:
                        pass
            if event is "Uninstall":
                # look for install script
                script = '-uninstall.cmd'
                exe = '_uninstall.exe'
                # print(path)
                listdir = os.listdir(path)
                fileList = []
                found = False

                for item in range(len(listdir)):
                    fileList.append(listdir.pop().lower())

                for file in fileList:
                    # print(file)
                    if script in file:
                        # print("\nInstall Script found.")
                        # run script directly
                        try:
                            os.startfile(path + '\\' + file)
                        except:
                            pass

                        found = True
                        break
                    elif exe in file:
                        # print("\nInstall Executable found.")
                        # run script directly
                        try:
                            os.startfile(path + '\\' + file)
                        except:
                            pass
                        # print("\n")
                        found = True
                        break
                if found == False:
                    # open directory
                    # print("Manual Install Required.\n")
                    try:
                        os.startfile(path)
                    except:
                        pass
            if event is "Open Software Library":
                try:
                    os.startfile("http://142.174.118.47/swlibrary/dbsearch.asp?search=" + sw4)
                except:
                    pass
            if event == 'Quit' or values is None:
                sys.exit()
            if event is "Next":
                window.close()
                return 0

        return


def iter(i, score, visited):
    if score[i] in visited:
        score[i] = score[i] - 1
        iter(i, score, visited)
    else:
        visited.append(score[i])



def installApps(apps):
    os.system('color 0A')

    for line in apps:
        if len(line) > 1:
            try:
                win2.close()
                del win2
            except:
                pass
            layout = [[sg.Listbox(apps, font=("Helvetica", 14), select_mode='multiple', size=(45, 8))]]
            win2 = sg.Window('Applications to Install', layout, location=(0, 0), default_element_size=(40, 1),
                             grab_anywhere=False).Finalize()
            sw_search(line)

        # print("\n--------------------------------------------------------------------------\n")

def append_corpads(assetname):
    return (assetname + '.corp.ads')

def main(*old):
    # print("Loading...")

    # super spaghetti :')

    while (True):
        newAsset = getAsset()

        # opens files for writing
        # file1 = open("full_list_" + getpass.getuser() + ".txt", "w")
        file2 = open("C:\TEMP\SOFTWARE_" + getpass.getuser() + ".txt", "w")

        # unfiltered list of all software
        full_list = []
        # filtered list that actually displays in window
        appsList = []

        try:
            full_list = getkeys(newAsset, full_list, win32con.HKEY_LOCAL_MACHINE, win32con.KEY_WOW64_32KEY)
        except FileNotFoundError:
            pass

        try:
            full_list = getkeys(newAsset, full_list, win32con.HKEY_LOCAL_MACHINE, win32con.KEY_WOW64_64KEY)
        except FileNotFoundError:
            pass

        try:
            full_list = getkeys(newAsset, full_list, win32con.HKEY_CURRENT_USER, 0)
        except FileNotFoundError:
            pass

        # fileWrite(full_list, file1)
        # file1.close()
        if len(full_list) == 0:
            layout = [[sg.T('Cannot connect to ' + newAsset +'. \n \n Try direct IP address.')],
                      [sg.OK(size=(18, 3), font='Arial 12')]]

            window = sg.Window('Error', layout)
            window.Read()
            window.Close()
            sys.exit()

        elif len(full_list) > 0:
            del full_list[0]  # first line always blank

        # print("Please make software selection.")

        # GUI window
        appsList = getApps(full_list)

        fileWrite(appsList, file2)

        file2.close()

        # print("success")
        break

    # time to install selected applications
    installApps(appsList)

    layout = [[sg.T('All done. \n \nMade by Bob Tian :)')],
              [sg.OK(size=(18, 3), font='Arial 12')]]

    window = sg.Window('Completed', layout)

    event, values = window.Read()
    window.Close()

    os.remove("pagehtml.txt")

    sys.exit()


if __name__ == "__main__":
    main()
