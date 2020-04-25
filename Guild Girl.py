import wmi
import socket
from gtts import gTTS
import os
import sys
import uuid 
import contextlib
with contextlib.redirect_stdout(None):
    import pygame as pg

#--------------------------------------------------------------------------------------------
#
#Created by P.D.M.Dilan
#2020/04/24

#--------------------------------------------------------------------------------------------

banner=(r"""

                     ██████╗ ██╗   ██╗██╗██╗     ██████╗ 
                    ██╔════╝ ██║   ██║██║██║     ██╔══██╗
                    ██║  ███╗██║   ██║██║██║     ██║  ██║
                    ██║   ██║██║   ██║██║██║     ██║  ██║
                    ╚██████╔╝╚██████╔╝██║███████╗██████╔╝
                     ╚═════╝  ╚═════╝ ╚═╝╚══════╝╚═════╝ 
                             ██████╗ ██╗██████╗ ██╗      
                            ██╔════╝ ██║██╔══██╗██║      
                            ██║  ███╗██║██████╔╝██║      
                            ██║   ██║██║██╔══██╗██║      
                            ╚██████╔╝██║██║  ██║███████╗ 
                             ╚═════╝ ╚═╝╚═╝  ╚═╝╚══════╝ 
                                                         
          ~~  Vocal Utility For Microsoft Windows System Informations ~~
                          ~~ Guild Girl version 0.0 ~~
                
                """)

print(banner)
print('\n')


#Gathering windows system infromation

computer = wmi.WMI()

os_info = computer.Win32_OperatingSystem()[0]
proc_info = computer.Win32_Processor()[0]
gpu_info = computer.Win32_VideoController()[0]


os_name = os_info.Name.encode('utf-8').split(b'|')[0]
os_version = ' '.join([os_info.Version, os_info.BuildNumber])
system_ram = round(float(os_info.TotalVisibleMemorySize) / 1048576 ) # KB to GB

#--------------------------------------------------------------------------------------------

print('Gathering system infromations \n')


#get the default download path
#------------------------------------------------------------------------------------------------------------------
def get_download_path():
#"""Returns the default downloads path for linux or windows"""
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')


downloadPath=get_download_path()

#writing windows system infromation into a file
stdoutOrigin=sys.stdout 
sys.stdout = open(downloadPath+'\\'+'info.txt', 'w')

print('Greeting '+ os.getlogin())
print('OS Name: {0}'.format(os_name))
print(' OS Version: {0}'.format(os_version))
print(' CPU: {0}'.format(proc_info.Name))
print(' RAM: {0} GB'.format(system_ram))
print(' Graphics Card: {0}'.format(gpu_info.Name))


print('Disk informations')

for d in computer.Win32_LogicalDisk():
    if d.Size is None or d.FreeSpace is None:
        print(d.Caption , d.FreeSpace, d.Size, d.Description)
    else:
        print(d.Caption , 'Disk size '+ str(round(float(d.Size) / 1073741824))+'GB', 'Free disk space '+str(round(float(d.FreeSpace)/ 1073741824))+'GB', d.Description)

hostname = socket.gethostname()    
IPAddr = socket.gethostbyname(hostname)    
print(" Hostname :" + hostname)    
print(" IP address :" + IPAddr) 

print (" MAC address  ", end="")
print (':'.join(['{:02x}'.format((uuid.getnode() >> ele) & 0xff) for ele in range(0,8*6,8)][::-1])) 


sys.stdout.close()
sys.stdout=stdoutOrigin

#end of writing information
#------------------------------------------------------------------------------------------------------------------

#converting text into speach

sysInfo=open(downloadPath+'\\'+'info.txt', 'r')
readTxt=sysInfo.read()

language='en-us'

try:
    oFile=gTTS(text=readTxt,lang=language,slow=False)
    oFile.save(downloadPath+'\\'+'sysInfo.mp3')
except:
    print('Check your internet connection !')
    print('Still you can access Info.txt \n')
    os.remove(downloadPath+'\\'+'sysInfo.mp3')
    exit()


print('Infromation gathering completed ! ')

#------------------------------------------------------------------------------------------------------------------

# playing the audio file

music_file=downloadPath+'\\'+'sysInfo.mp3'

pg.mixer.init()
pg.mixer.music.set_volume(0.6)
clock = pg.time.Clock()

pg.mixer.music.load(music_file)
print("You can find info files at {} ".format(music_file) +' and sysInfo.txt \n')

pg.mixer.music.play()
while pg.mixer.music.get_busy():
    # check if playback has finished
    clock.tick(30)

#------------------------------------------------------------------------------------------------------------------
