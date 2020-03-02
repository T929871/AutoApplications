import win32con
import winreg
import os
import shutil
import unicodedata
import time
import PySimpleGUI as sg
import getpass
import sys

import urllib.request
import shutil
import difflib
import re
from fuzzywuzzy import fuzz
from fuzzywuzzy import process



def getAsset():
    layout = [[sg.Text('Enter old Asset #')],      
                     [sg.InputText()],      
                     [sg.Submit(), sg.Cancel()]]      

    window = sg.Window('Auto Applications Tool', layout)    

    event, text_input = window.Read()    
    window.Close()

    asset = text_input  
    #print(asset)
    return asset[0]


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
            
            #preliminary filtering
            ignoreList = ['runtime', 'redistributable', 'installer', 'setting', 'updater', 'update for ', 'extension', 'proofing', 'flash player', 'add-in', 'driver', 'activex', 'anyconnect', 'microsoft access', 'microsoft dcf', 'microsoft excel', 'microsoft groove', 'infopath', 'lync', 'components', 'office osm', 'office shared', 'onenote', 'outlook', 'powerpoint', 'microsoft word', 'parser', 'myid', 'publisher', 'silverlight', 'framework', 'policy', 'plug-in', 'help', 'webcam', 'driveguard', 'service pack', 'support', 'microsoft office', 'hp ', 'intel ', 'intel(r) ', 'apple ', 'viewer'] #lowercase
            skip = False
            for item in ignoreList:
                if item in val.lower():
                    skip = True
            if skip == False:
                full_list.append(val)            
            else:
                continue
                #skip 
                 
        except EnvironmentError:
            continue
    
    full_list.sort() #alpha sort
    
    return full_list


def getApps(full_list):
    layout = [
        [sg.Listbox(full_list, select_mode='multiple', size=(70, 20))],
        #[sg.Button("Select All")], [sg.Button("Clear Selection")], #not really that useful
        [sg.Submit(tooltip='Click to submit this window'), sg.Cancel()]         
         
    ]                 

    window = sg.Window('Select Applications to Install', layout, default_element_size=(40, 1), grab_anywhere=False)

    while True:      
        event, list = window.Read()      
        if event is None or event == 'Exit' or event == 'Cancel':      
            sys.exit()

        #if none selected
        elif event == 'Submit' and not list[0]:
            
            layout = [[sg.Text('No Applications Selected!')],[sg.Cancel()]] 
            pop = sg.Window(':O', layout)
            pop.Read()
            pop.Close()
          
        else:                    
            print(list[0])
            window.Close()
            return list[0]           

    window.Close()



#
def fileWrite(list, file):
    for item in list:
        file.write(item+'\r\n')
    return
    
#search software from library    
def sw_search(software):

    print("\nLooking for: \n" + software)

   
    # get first 4 letters
    sw4 = software[:4].lower()
    # remove spaces
    sw4 = sw4.replace(" ", "")
     
    # special cases *********************************************************************************************************
    if sw4 == 'micr':
        sw4 = 'msft'

    print('Search token \"' + sw4 + '"\n')
     
    urlstr = "http://142.174.118.47/swlibrary/dbsearch.asp?search=" + sw4
    # print(urlstr)
    
    # open page
    #os.startfile(urlstr)

    #
    parsePage(urlstr, software, sw4)
    
# turn page into text for processing
def parsePage(url, software, sw4):
    pagehtml = open("pagehtml.txt", "wb")
    
    page = urllib.request.urlopen(url)
    shutil.copyfileobj(page, pagehtml)    

    with open('pagehtml.txt',encoding='cp1252') as f:
        data = f.read()
    with open('pagehtml.txt','w',encoding='utf8') as f:
        f.write(data)   
    pagehtml.close()
    
    readPage(software, sw4)


def readPage(software, sw4):
    with open('pagehtml.txt', encoding='utf8', errors='ignore') as f:
        pagehtml = f.read()
        #print(pagehtml)

# fine tuning here
#************************************************************************************************************************
        #SPECIAL CASES
        #---bluezone
        if "BlueZone " in software:
            software = software.replace('BlueZone ', 'BlueZone Desktop')
        
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
        #print(path)
        pathCmp = re.sub('\W+','', path)
 #************************************************************************************************************************       
        page_lines = pagehtml.splitlines()
               
        score = []
        cmper = 'null'
        for line in page_lines:
            cmper = re.sub('\W+','', line)
            score.append(fuzz.ratio(path, cmper))
            #print(line)
        index = score.index(max(score))
        
        html = page_lines[index]
        #print(html)
        
        #print()
        try:
            url = re.findall(r'file\:\/\/(.*?)"\>', html)
            path = url[0]
        except IndexError:
            print("*** Not in Software Library ***\n\n")
            return
        if "rosoftApps" in path[33:]:
            print("*** Not in Software Library ***\n\n")
            return
        
        print("Found: \n" + path[33:])
        print("\n--- Match confidence: " + str(score.index(max(score))) + " ---")
        
        choice = input("\n1. Auto install \n2. Open Software Library \n\n==> Enter 1 or 2: ")
        
        if choice == '1':
            # look for install script
            script = '-install.cmd'
            exe = '-install.exe'
            print(path)
            listdir = os.listdir(path)
            fileList = []
            found = False
            
            for item in range(len(listdir)):
                fileList.append(listdir.pop().lower())
                   
            for file in fileList:
                #print(file)
                if script in file:                
                    print("\nInstall Script found.")               
                    #run script directly
                    os.startfile(path+'\\'+file)
                    print("\n")
                    found = True
                    break
                elif exe in file:               
                    print("\nInstall Executable found.")               
                    #run script directly
                    os.startfile(path+'\\'+file)
                    print("\n")
                    found = True
                    break
                    
            if found == False:         
                # open directory
                print("Manual Install Required.\n")
                os.startfile(path)
                
        elif choice == '2':
            os.startfile("http://142.174.118.47/swlibrary/dbsearch.asp?search=" + sw4)
            print("\n")
         
        else:
            print("\nSkipped**********************************************************\n")

def main():
    print("Loading...")
    while(True):
        print("Please make software selection.")
        newAsset = getAsset()
        
        #opens files for writing
        file1 = open("full_list_" + getpass.getuser() + ".txt", "w")
        file2 = open("C:\TEMP\SOFTWARE_" + getpass.getuser() + ".txt", "w")
        
        #unfiltered list of all software
        full_list = []
        #filtered list that actually displays in window
        appsList = []
                
                
        full_list = getkeys(newAsset, full_list, win32con.HKEY_LOCAL_MACHINE, win32con.KEY_WOW64_32KEY)
        full_list = getkeys(newAsset, full_list, win32con.HKEY_LOCAL_MACHINE, win32con.KEY_WOW64_64KEY)
        try: #doesnt exist on windows 7
            full_list = getkeys(newAsset, full_list, win32con.HKEY_CURRENT_USER, 0)
        except FileNotFoundError:
            pass
        
        #GUI window
        appsList = getApps(full_list)
        
        fileWrite(full_list, file1)
        fileWrite(appsList, file2)
        
        file1.close()
        file2.close()
        
        print("success")
        break
        
    
    #time to install selected applications
    apps = appsList
    #change bg to blue
    os.system('color 0A')
    for line in apps:
        if len(line) > 1:
            sw_search(line)
            input("Press enter to continue...")
            print("\n--------------------------------------------------------------------------\n")
    
    print("All done. \nMade by Bob Tian :)")
    input()
    sys.exit()
    
if __name__ == "__main__":
    main()

