import os
import subprocess
import tkinter as tk
from tkinter.constants import BOTH, BUTT, LEFT, RAISED
import webbrowser
from tkinter import ttk
from tkinter import *
from PIL import ImageTk
from PIL import Image

global root
root=tk.Tk()

def mainmenu():
    #Frames Used
    global fr2
    fr2=tk.LabelFrame(root,text="Online KMS Activation",fg="white")
    global fr1
    fr1=tk.LabelFrame(root,text="Main Options",fg="White")
    global fr3
    fr3=tk.LabelFrame(root,text="Extras",fg="white")

    root.title("PyMAS 1.0")
    root.config(background="#16161a")
    root.resizable(False, False)
    Main_Options()
    menubars()
    root.config(menu=menubar)
    root.mainloop()        

def menubars():
    #Menubars Used
    global menubar
    menubar = tk.Menu(root,background="#16161a",fg="white")
    #Extras Menu
    Extras = tk.Menu(menubar, tearoff=False,bg="#16161a",fg="white")
    Extras.add_command(label="HWID Readme",command=lambda:(run("hwidr")),activebackground="#383842")
    Extras.add_command(label="KMS38 Readme",command=lambda:(run("kms38r")),activebackground="#383842")
    Extras.add_command(label="Online KMS Readme",command=lambda:(run("oldkmsr")),activebackground="#383842")
    Extras.add_separator()
    Extras.add_command(label="Extract OEM Readme",command=lambda:(run("extarctoemr")),activebackground="#383842")
    Extras.add_command(label="KMS38 Protection Readme",command=lambda:(run("kms38protr")),activebackground="#383842")
    menubar.add_cascade(label="ReadMe's", menu=Extras)
    #Help Menu
    helpmenu = tk.Menu(menubar, tearoff=False,bg="#16161a",fg="white")
    helpmenu.add_command(label="Credits",activebackground="#383842",command=lambda:(run("credits")))
    helpmenu.add_command(label="About Creator",activebackground="#383842",command=githubme)

    menubar.add_cascade(label="Help",menu=helpmenu)

def githubme():
    webbrowser.open("https://github.com/geoseiden")

def run(a):
    if a=="check":
        subprocess.call([r"PyMAS_Files\Check_Activation.cmd"])
    elif a=="hwidact":
        subprocess.call([r"PyMAS_Files\HWID_Activation.cmd"])
    elif a=="kms38act":
        subprocess.call([r"PyMAS_Files\KMS38_Activation.cmd"])
    elif a=="onlkms":
        onlkmsact()
    elif a=="oldkmsact":
        subprocess.call([r"PyMAS_Files\Old-act\Activate.cmd"])
    elif a=="oldkmsren":
        subprocess.call([r"PyMAS_Files\Old-act\Renewal_Setup.cmd"])
    elif a=="oldkmsunin":
        subprocess.call([r"PyMAS_Files\Old-act\Uninstall.cmd"])
    elif a=="hwidr":
        os.system(r"notepad.exe PyMAS_Files\Readmefiles\hwid_readme.txt")
    elif a=="kms38r":
        os.system(r"notepad.exe PyMAS_Files\Readmefiles\kms38_readme.txt")
    elif a=="extractoemr":
        os.system(r"notepad.exe PyMAS_Files\Readmefiles\extractoem_readme.txt")
    elif a=="kms38prot":
        os.system(r"notepad.exe PyMAS_Files\Readmefiles\kms38prot_readme.txt")
    elif a=="oldkmsr":
        os.system(r"notepad.exe PyMAS_Files\Readmefiles\oldact_readme.txt")
    elif a=="exOEM":
        subprocess.call([r"PyMAS_Files\Extras\Extract_OEM_Folder.cmd"])
    elif a=="win10install":
        subprocess.call([r"PyMAS_Files\Extras\OEMRET-Install_W10_Key.cmd"])
    elif a=="win10change":
        subprocess.call([r"PyMAS_Files\Extras\OEMRET-Change_W10_Edition.cmd"])
    elif a=="kms38prot":
        subprocess.call([r"PyMAS_Files\Extras\Protect_Unprotect-KMS38.cmd"])
    elif a=="abbodi":
        webbrowser.open("https://forums.mydigitallife.net/threads/74197/")
    elif a=="credits":
        os.system(r"notepad.exe PyMAS_Files\Readmefiles\Credits.txt")
    elif a=="isodown":
        webbrowser.open("https://www.heidoc.net/php/Windows-ISO-Downloader.exe")

def onlkmsact():
    fr1.destroy()

    fr2.grid(row=0,column=0)
    fr2.configure(bg="#16161a")

    l1=tk.Label(fr2,text="This will skip any permanent HWID and KMS38 Activation.",bg="#16161a",fg="white")
    l1.grid(row=1,column=0)
    
    l2=tk.Label(fr2,text="KMS activates for 180 Days.(For core/ProWMC edition it is 30/45 Days)",bg="#16161a",fg="white")
    l2.grid(row=2,column=0)
    
    l3=tk.Label(fr2,text="_____________________________________________________________________________________________",bg="#16161a",fg="#16161a")
    l3.grid(row=3,column=0)

    b1=tk.Button(fr2,text="Activate - Windows / Server / Office",command=lambda:run("oldkmsact"),borderwidth=0,bg="#7f5af0",fg="white")
    b1.grid(row=4,column=0)
    coh(b1,"#616161","#7f5af0")
    
    l4=tk.Label(fr2,text="_______________________________",bg="#16161a",fg="#16161a")
    l4.grid(row=5,column=0)

    b2=tk.Button(fr2,text="Activation Renewal",command=lambda:run("oldkmsren"),borderwidth=0,bg="#7f5af0",fg="white")
    b2.grid(row=6,column=0)
    coh(b2,"#616161","#7f5af0")

    l5=tk.Label(fr2,text="_______________________________",bg="#16161a",fg="#16161a")
    l5.grid(row=7,column=0)

    b3=tk.Button(fr2,text="Complete Uninstall",command=lambda:run("oldkmsunin"),borderwidth=0,bg="#7f5af0",fg="white")
    b3.grid(row=8,column=0)
    coh(b3,"#616161","#7f5af0")

    l6=tk.Label(fr2,text="_______________________________",bg="#16161a",fg="#16161a")
    l6.grid(row=9,column=0)

    b4=tk.Button(fr2,text="Back",command=lambda:(fr2.destroy(),mainmenu()),borderwidth=0,bg="red",fg="white")
    b4.grid(row=10,column=0,sticky="w")
    coh(b4,"#616161","red")

def extras():
    fr1.destroy()

    fr3.grid(row=0,column=0)
    fr3.configure(bg="#16161a")

    l1=tk.Label(fr3,text="_____________________________________________________________________________________________",bg="#16161a",fg="#16161a")
    l1.grid(row=1,column=0)

    b1=tk.Button(fr3,text="Extract $OEM$ Folder [Preactivation]",command=lambda:(run("exoem")),borderwidth=0,bg="#7f5af0",fg="white")
    b1.grid(row=2,column=0)
    coh(b1,"#616161","#7f5af0")

    l2=tk.Label(fr3,text="_______________________________",bg="#16161a",fg="#16161a")
    l2.grid(row=3,column=0)

    b2=tk.Button(fr3,text="Insert Windows 10 Key [OEMRET]",command=lambda:(run("win10install")),borderwidth=0,bg="#7f5af0",fg="white")
    b2.grid(row=4,column=0)
    coh(b2,"#616161","#7f5af0")

    l3=tk.Label(fr3,text="_______________________________",bg="#16161a",fg="#16161a")
    l3.grid(row=5,column=0)

    b3=tk.Button(fr3,text="Change Windows 10 Edition [OEMRET]",command=lambda:(run("win10change")),borderwidth=0,bg="#7f5af0",fg="white")
    b3.grid(row=6,column=0)
    coh(b3,"#616161","#7f5af0")

    l4=tk.Label(fr3,text="_______________________________",bg="#16161a",fg="#16161a")
    l4.grid(row=7,column=0)

    b4=tk.Button(fr3,text="Protect / Unprotect KMS38 Activation",command=lambda:(run("kms38prot")),borderwidth=0,bg="#7f5af0",fg="white")
    b4.grid(row=8,column=0)
    coh(b4,"#616161","#7f5af0")

    l5=tk.Label(fr3,text="_______________________________",bg="#16161a",fg="#16161a")
    l5.grid(row=9,column=0)

    b5=tk.Button(fr3,text="Download Windows ISO Downloader",command=lambda:(run("isodown")),borderwidth=0,bg="#7f5af0",fg="white")
    b5.grid(row=10,column=0)
    coh(b5,"#616161","#7f5af0")

    l6=tk.Label(fr3,text="_______________________________",bg="#16161a",fg="#16161a")
    l6.grid(row=9,column=0)

    b6=tk.Button(fr3,text="abbodi1406's Batch Scripts Repo [Link]",command=lambda:(run("abbodi")),borderwidth=0,bg="#7f5af0",fg="white")
    b6.grid(row=10,column=0)
    coh(b6,"#616161","#7f5af0")
    
    l7=tk.Label(fr3,text="_______________________________",bg="#16161a",fg="#16161a")
    l7.grid(row=11,column=0)

    b7=tk.Button(fr3,text="Back",command=lambda:(fr3.destroy(),mainmenu()),borderwidth=0,bg="red",fg="white")
    b7.grid(row=12,column=0,sticky="w")
    coh(b7,"#616161","red")
    
def coh(button, colorOnHover, colorOnLeave):

    button.bind("<Enter>", func=lambda e: button.config(
        background=colorOnHover))
  
    button.bind("<Leave>", func=lambda e: button.config(
        background=colorOnLeave))

def Main_Options():

        fr1.grid(row=0,column=0)
        fr1.configure(bg="#16161a")

        b1=tk.Button(fr1,text="Check Activation",command=lambda:(root.destroy(),run("check")),borderwidth=0,bg="#7f5af0",fg="white")
        b1.grid(row=0,column=0)
        coh(b1,"#616161","#7f5af0")
        
        l1=tk.Label(fr1,text="_____________________________________________________________________________________________",bg="#16161a",fg="#16161a")
        l1.grid(row=1,column=0)
        
        b2=tk.Button(fr1,text="HWID Activation",command=lambda:(root.destroy(),run("hwidact")),borderwidth=0,bg="#7f5af0",fg="white")
        b2.grid(row=2,column=0)
        coh(b2,"#616161","#7f5af0")

        l2=tk.Label(fr1,text="_______________________________",bg="#16161a",fg="#16161a")
        l2.grid(row=3,column=0)

        b3=tk.Button(fr1,text="KMS38 Activation",command=lambda:(root.destroy(),run("kms38act")),borderwidth=0,bg="#7f5af0",fg="white")
        b3.grid(row=4,column=0)
        coh(b3,"#616161","#7f5af0")

        l3=tk.Label(fr1,text="_______________________________",bg="#16161a",fg="#16161a")
        l3.grid(row=5,column=0)

        b4=tk.Button(fr1,text="Online KMS Activation",command=lambda:(onlkmsact()),borderwidth=0,bg="#7f5af0",fg="white")
        b4.grid(row=6,column=0)
        coh(b4,"#616161","#7f5af0")

        l4=tk.Label(fr1,text="_______________________________",bg="#16161a",fg="#16161a")
        l4.grid(row=7,column=0)

        b5=tk.Button(fr1,text="Extras",command=lambda:(extras()),borderwidth=0,bg="#7f5af0",fg="white")
        b5.grid(row=8,column=0)
        coh(b5,"#616161","#7f5af0")


mainmenu()