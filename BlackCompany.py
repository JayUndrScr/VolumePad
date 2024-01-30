#!/usr/bin/env python3.10.11
import os

import PIL.Image

#First try to import all requirements. if succesfull start application, else download all requirements via requirements file then import
try:
    import tkinter as tk
    from tkinter import ttk
    from PIL import Image,ImageTk
    import pyfirmata
    from pyfirmata import ArduinoNano, util, STRING_DATA
    from ctypes import cast, POINTER
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
    import serial.tools.list_ports
    import math
    import time
    import psutil, win32process, win32gui
    import subprocess

except:

    exe = 'buildLALArequirements.txt'

    #if we need find it first
    for root, dirs, files in os.walk(r'C:\Users'):
        for name in files:
            if name == exe:
                reqLink = os.path.abspath(os.path.join(root, name))
                print(reqLink)

    os.system(r'cmd /c "py -m pip install -r %s"' % (reqLink))

    import tkinter as tk
    from tkinter import ttk
    from PIL import Image, ImageTk
    import pyfirmata
    from pyfirmata import ArduinoNano, util, STRING_DATA
    from ctypes import cast, POINTER
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
    import serial.tools.list_ports
    import math
    import time
    import psutil, win32process, win32gui
    import subprocess
#===========================================================================
#Splash Screen
spl = tk.Tk()
spl.title("VPStudio")
spl.minsize(660, 200)
spl.maxsize(660, 200)
spl.configure(background='grey20')
splImg = Image.open(r"Images\vpStudio3.png")
splImg = splImg.resize((800,200), Image.Resampling.LANCZOS)
splResize = ImageTk.PhotoImage(splImg)
splScreen = tk.Canvas(spl, bg="grey20", width=500,height=130, highlightbackground="grey20")
splScreen.place(x=80,y=25)
splScreen.create_image((290,70), image=splResize)
splMes = tk.Label(spl, bg="grey20", fg="white", text="Loading...", font=('Arial', 20))
splMes.place(x=10, y=160)
x = (spl.winfo_screenwidth()//2)-(660//2)
y = (spl.winfo_screenheight() // 2) - (200 // 2)
spl.geometry('{}x{}+{}+{}'.format(660,200,x,y))
spl.overrideredirect(True)
spl.update()
#===========================================================================
#Variables
loopTheLoop = True
errorLoop = True
something = os.getcwd()
appPrint = []
appRange = 7
bar = []
appChoosen = []
oldValue = [-1,-1,-1,-1,-1,-1,-1]
value = [2.2,2,2,2,2,2,2]
storedValue = [1,1,1,1,1,1,1]
updateValue = [0,0,0,0,0,0,0]
#Methode to stop application
def stopRun():
    global loopTheLoop
    loopTheLoop = False
    win.destroy()
    board.exit()
#Methode to restart application
def rerun():
    global loopTheLoop
    loopTheLoop = False
    win.destroy()
    board.exit()
    os.startfile("SubDomain.exe")
def msg(text):
    if text:
        board.send_sysex(STRING_DATA, util.str_to_two_byte_iter(text))
def active_window_process_name():
    pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow()) #This produces a list of PIDs active window relates to
    try:
        print(psutil.Process(pid[-1]).name())
    except:
        pass
#Methode to get access to connected ports on device
def get_ports():
    ports = serial.tools.list_ports.comports()

    return ports
#Methode to find port of the given name
def findArduino(portsFound):
    commPort = 'None'
    numConnection = len(portsFound)

    for i in range(0, numConnection):
        port = foundPorts[i]
        strPort = str(port)

        if 'CH340' in strPort:
            splitPort = strPort.split(' ')
            commPort = (splitPort[0])

    return commPort
#==========================================================================
#Arduino Setup
foundPorts = get_ports()
connectPort = findArduino(foundPorts)
try:
    board = ArduinoNano(connectPort)
    for i in range(8):
        board.analog[i].mode = pyfirmata.INPUT
    it = pyfirmata.util.Iterator(board)
    it.start()
except:
    errorLoop = False
    spl.destroy()
    err = tk.Tk()
    err.title("VPStudio")
    err.minsize(660, 200)
    err.maxsize(660, 200)
    err.configure(background='firebrick')
    #startBtn = tk.Button(err, bd=10, bg="white", justify="right", text="Stop", command=stopErr)
    #startBtn.place(x=600, y=150)
    errMes = tk.Label(err, bg='grey90', text=" ERROR: USB not connected or COM port not detected", font=('Arial', 15))
    errMes.place(x=80, y=90)
    x = (err.winfo_screenwidth()//2)-(660//2)
    y = (err.winfo_screenheight() // 2) - (200 // 2)
    err.geometry('{}x{}+{}+{}'.format(660,200,x,y))
    err.overrideredirect(True)
    errCount = tk.Label(err,bg='firebrick', text="Closing in: 3", font=('Arial', 15))
    errCount.place(x=1, y=170)
    err.update()
    time.sleep(1.0)
    try:
        errCount.config(text="Closing in: 2")
        err.update()
        time.sleep(1.0)
    except:
        pass
    try:
        errCount.config(text="Closing in: 1")
        err.update()
        time.sleep(1.0)
    except:
        pass
    err.destroy()
    err.mainloop()
#=====================================================================
#Get acces to system volume mixer
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
sessions = AudioUtilities.GetAllSessions()
#function to get application currently in use
def apps():
    appList = []
    x = "none"
    for session in sessions:
        if session.Process:
            appList.append(session.Process.name())
    appList = list(set(appList))
    length = len(appList)
    for i in range(length):
        x = appList[i].replace(".exe","")
        appList.append(x)
    for i in range(length):
        appList.pop(0)
    #appList.append("Focused")
    appList.append("None")
    return appList
#=================================================================
#UI SETUP
#==================================================================
try:
    spl.destroy()
except:
    pass
msg("a2k")
#msg('f2')
win = tk.Tk()
win.title("VPStudio")
win.minsize(700,450)
win.maxsize(700,450)
win.configure(background='grey20')
titleLogo = Image.open(r"Images\LogoCMIT.ico")
titleLogo = ImageTk.PhotoImage(titleLogo)
win.iconphoto(False,titleLogo)
#win.withdraw()
#===================================================================
#Labels
image1 = Image.open(r"Images\vpStudio3.png")
image1 = image1.resize((200,50), Image.Resampling.LANCZOS)
logoResize = ImageTk.PhotoImage(image1)
Logo = tk.Canvas(win, bg="grey20", width=160,height=50, highlightbackground="grey20")
Logo.place(x=1,y=1)
Logo.create_image((80,25), image=logoResize)
#profileLabel = tk.Label(win, text="Profile", width=5, font=('Arial', 20), background="grey20", fg="white")
#profileLabel.place(x=320,y=10)
image2 = Image.open(r"Images\MasterVolumeIcon.png")
MIResize = ImageTk.PhotoImage(image2)
MI = tk.Canvas(win, bg="grey20", width=32,height=32, highlightbackground="grey20")
MI.grid(column=1,row=2,padx=(0,50),pady=(60,20))
MI.create_image((18,18), image=MIResize)
#masterLabel = tk.Label(win, text=" Master:", width=5, font=('Arial', 20), background="grey20", fg="white")
#masterLabel.grid(column=1,row=2,padx=(50,50),pady=(60,20))
masterPrint = tk.Label(win, text=(""), width=2, font=('Arial', 10),background="grey20", fg="white")
masterPrint.grid(column=1,row=5,padx=(0,50))
for i in range(appRange):
    nue = tk.Label(win, text=(""), width=2, font=('Arial', 10), background="grey20", fg="white")
    nue.grid(column=2+i,row=5)
    appPrint.append(nue)
#==================================================================
#ProgressBars
barS = ttk.Style()
barS.theme_use("clam")
barS.configure("purple.Vertical.TProgressbar",troughcolor="grey15", lightcolor="#9225ff", darkcolor="#9225ff", background='#9225ff',bordercolor="grey20")
barMaster = ttk.Progressbar(win,style="purple.Vertical.TProgressbar", orient="vertical", mode="determinate", length=250)
barMaster.grid(column=1, ipadx= 3, row=4,padx=(50,100))
for i in range(appRange):
    labda = ttk.Progressbar(win,style="purple.Vertical.TProgressbar", orient="vertical", mode="determinate", length=250)
    labda.grid(column=2+i, ipadx= 3, row=4,)
    bar.append(labda)
#==================================================================
#IconMenu
n = [8,8,8,8,8,8,8]
es = [9,9,9,9,9,9,9]
noImg = Image.open('Images\\None.png')
noImg = ImageTk.PhotoImage(noImg)
for i in range(appRange):
    n[i] = tk.StringVar()
    n[i].set("None")
    es[i] = tk.OptionMenu(win, n[i], *apps())
    es[i].configure(indicatoron=False, fg="grey20",highlightthickness=0, bg="grey20",activebackground='grey20', image=noImg,  borderwidth=0)
    es[i].grid(column=2+i, row=2,pady=(60,20), padx=(10,10))

oldApp = [n[0].get(),n[1].get(),n[2].get(),n[3].get(),n[4].get(),n[5].get(),n[6].get()]
doubleCheck = ['x','x','x','x','x','x','x']
def changeIcon(index):
    for i in range(len(apps())):
        if(n[index].get() == apps()[i] and n[index].get() in doubleCheck):
            try:
                appImg = Image.open("Images\\none.png")
                newAppImg = Image.open("Images\\"+apps()[i] + ".png")
                indexFound = doubleCheck.index(n[index].get())
                n[indexFound].set('None')
                doubleCheck[indexFound] = 'x'
                if (n[index].get() == 'None'):
                    doubleCheck[index] = 'x'
                else:
                    doubleCheck[index] = n[index].get()
                oldApp[indexFound] = 'None'
                oldApp[index] = n[index].get()
            except:
                appImg = Image.open("Images\\none.png")
                newAppImg = Image.open("Images\\question.png")
                indexFound = doubleCheck.index(n[index].get())
                n[indexFound].set('None')
                doubleCheck[indexFound] = 'x'
                if (n[index].get() == 'None'):
                    doubleCheck[index] = 'x'
                else:
                    doubleCheck[index] = n[index].get()
                oldApp[indexFound] = 'None'
                oldApp[index] = n[index].get()
            #for i in range(appRange):
            #    print(n[i].get())
            #print(doubleCheck)
            newAppLabel = tk.Label(win)
            newAppLabel.image = ImageTk.PhotoImage(newAppImg)
            newFilePath = '\\examplePics\\' + n[index].get() + ".png"
            os.chdir(os.getcwd() + r'\imageToScreen\source')
            try:
                subprocess.check_call(
                    os.getcwd() + r'\imageToScreen.exe -f "' + os.getcwd() + newFilePath + '" -s ' + str(index) + ' -R -q')
            except Exception as e:
                print(e)
            with open(os.getcwd() + '\\hexFiles\\' + str(index) + ".txt") as file:
                msg("s" + str(6 - index))

                msg("i")
                content = file.read().strip()
                single = content.split(", ")
                #print(len(content))
                #print(len(single))
                for i in (range(int(len(single) / 4))):
                    final = []
                    for j in range(4):
                        # print(str(i) + " " + str(j))
                        final.append(single[i * 4 + j])
                    #print(",".join(final))
                    msg(",".join(final))
                msg("p")
                msg("o")
                os.chdir(something)
            es[index].config(image=newAppLabel.image)
            appLabel = tk.Label(win)
            appLabel.image = ImageTk.PhotoImage(appImg)
            msg("u" + str(6 - indexFound))
            os.chdir(something)
            es[indexFound].config(image=appLabel.image)
        elif(n[index].get() == apps()[i]):
            try:
                appImg = Image.open("Images\\"+ apps()[i] + ".png")
                oldApp[index] = n[index].get()
                if(n[index].get() == 'None'):
                    doubleCheck[index] = 'x'
                else:
                    doubleCheck[index] = n[index].get()
                #print(doubleCheck)
            except:
                appImg = Image.open("Images\\question.png")
                oldApp[index] = n[index].get()
                if (n[index].get() == 'None'):
                    doubleCheck[index] = 'x'
                else:
                    doubleCheck[index] = n[index].get()
                #print(doubleCheck)
            appLabel = tk.Label(win)
            appLabel.image = ImageTk.PhotoImage(appImg)
            es[index].config(image=appLabel.image)
            filePath = '\\examplePics\\' + n[index].get() +".png"
            os.chdir(os.getcwd() + r'\imageToScreen\source')
            try:
                subprocess.check_call(
                    os.getcwd() + r'\imageToScreen.exe -f "' + os.getcwd() + filePath + '" -s '+ str(index) +' -R -q')
            except Exception as e:
                print(e)
            with open(os.getcwd() + '\\hexFiles\\' + str(index) + ".txt") as file:
                msg("s" +str(6 - index))

                msg("i")
                content = file.read().strip()
                single = content.split(", ")
                #print(len(content))
                #print(len(single))
                for i in (range(int(len(single) / 4))):
                    final = []
                    for j in range(4):
                        # print(str(i) + " " + str(j))
                        final.append(single[i * 4 + j])
                    #print(",".join(final))
                    msg(",".join(final))
                msg("p")
                msg("o")
                os.chdir(something)
#=====================================================================
#Buttons
# startBtn = tk.Button(win, bd=10, bg="#FF0000", justify= "right", text="Stop", command=stopRun)
# startBtn.place(x=550,y=400)
refreshBtn = tk.Button(win, bd=10, bg="#87CEEB", justify= "right", text="Refresh", command=rerun)
refreshBtn.place(x=600,y=400)
#======================================================================
#=======================================================================
#Current status of Mastervolume
mOlded= -1
volume_percentage = 6
def checkMaster(valueM):
    global volume_percentage
    volume_percentage = round((1 - valueM) * 100)
    mid = round(volume_percentage)
    masterPrint.config(text=str(mid), fg="white")
    if volume_percentage == 11:
        dB = -32
    elif volume_percentage == 10:
        dB = -33
    elif volume_percentage == 9:
        dB = -35
    elif volume_percentage == 8:
        dB = -37
    elif volume_percentage == 7:
        dB = -38
    elif volume_percentage == 6:
        dB = -39
    elif volume_percentage == 5:
        dB = -41
    elif volume_percentage == 4:
        dB = -44
    elif volume_percentage == 3:
        dB = -49
    elif volume_percentage == 2:
        dB = -54
    elif volume_percentage == 1:
        dB = -60
    elif volume_percentage == 0:
        dB = -65
    else:
        dB = (34 * math.log(volume_percentage / 100, 10))
    #print(volume_percentage)
    try:
        volume.SetMasterVolumeLevel(dB, None)
    except:
        pass
#Current status of Application audio sources
def checkApp(index, valueUpdate):
        for session in sessions:
            global value
            volume1 = session._ctl.QueryInterface(ISimpleAudioVolume)
            value[index] = math.floor((1 - valueUpdate) * 100) /100
            dB = value[index] * 100
            appPrint[index].config(text=str(int(dB)), fg="white")
            if session.Process and session.Process.name() == n[index].get()+ ".exe":
                volume1.SetMasterVolume(value[index], None)
#Protocol for pressing close window button(top right)
win.protocol("WM_DELETE_WINDOW", stopRun)
#MainLoop
while loopTheLoop and errorLoop:

    mValue = board.analog[0].read()
    if(mOlded > (mValue + 0.003) or mOlded < (mValue - 0.003)):
        checkMaster(mValue)
        mOlded = mValue

    for a in range(appRange):
        updateValue[a] = board.analog[1 + a].read()
        if(oldValue[a] > (updateValue[a] + 0.003) or oldValue[a] < (updateValue[a] - 0.003)):
            checkApp(a, updateValue[a])
            oldValue[a] = updateValue[a]
    for i in range(appRange):
        barMaster['value'] = masterPrint['text']
        bar[i]['value'] = appPrint[i]['text']

    for q in range(appRange):
        if(n[q].get()!=oldApp[q]):
            checkApp(q, board.analog[1 + q].read())
            changeIcon(q)
    win.update()
    # p = psutil.Process(20500)
    # print(p.exe())
    time.sleep(0.01)