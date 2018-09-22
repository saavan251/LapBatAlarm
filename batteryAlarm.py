import psutil
import os
import subprocess

from time import time, sleep

vbsScript = '''
set speech = Wscript.CreateObject("SAPI.spVoice") 
speech.speak '''
newMsg = ''

batteryStatus = psutil._common.sbattery._asdict(psutil.sensors_battery())
#print(batteryStatus)
cwd = "C:\\Users\\VISHAL ASHANK\\Desktop\\projects\\Scripts\\batteryAlarm\\"
if(batteryStatus['percent']>95):
    vbsScript += '\"your battery is charged to ' + str(batteryStatus['percent']) + ' percent please turn off power supply to your laptop.\"'
    fileName = cwd + '$temp' + '.vbs'
    file = open(fileName,'w')
    file.write(vbsScript)
    file.close()
    #print(fileName)
    for i in range(0,2):
        sleep(2)
        status = psutil._common.sbattery._asdict(psutil.sensors_battery())
        if(status['power_plugged'] == True):
            process = subprocess.Popen(fileName,stdout= subprocess.PIPE,shell = True)
            process.communicate()
    os.remove(fileName)
