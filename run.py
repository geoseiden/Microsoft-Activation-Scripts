import os
import csv
import sys
import webbrowser
import subprocess
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox


def rp(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

global root
root=tk.Tk()
global bgc
bgc="#121212"
global fgc
fgc="White"
global btc
btc="#2B2B2B"
st=ttk.Style()
root.tk.call("source",rp("PyWinAct_Files/azure-dark.tcl"))
st.theme_use("azure-dark")

def mainmenu():
    root.iconbitmap(rp("PyWinAct_Files\pm.ico"))
    #Frames Used
    global fr2
    fr2=tk.LabelFrame(root,text="Online KMS Activation",fg="white")
    global fr1
    fr1=tk.LabelFrame(root,text="Main Options",fg="White")
    global fr3
    fr3=tk.LabelFrame(root,text="Extras",fg="white")
    global fr4
    fr4=tk.LabelFrame(root,text="Windows ISO Downloader",fg="white")
    global fr5
    fr5=tk.LabelFrame(root,text="Office ISO Downloader",fg="white")

    root.title("PyWinAct 2.0 By Geoseiden")
    root.config(background=bgc)
    root.resizable(False, False)
    Main_Options()
    menubars()
    root.config(menu=menubar)
    root.mainloop()        

def Main_Options():

        fr1.grid(row=0,column=0)
        fr1.configure(bg=bgc)

        b1=tk.Button(fr1,text="Check Activation",command=lambda:(root.destroy(),run("check")),borderwidth=0,bg=btc,fg="white",relief=RAISED)
        b1.grid(row=0,column=0)
        coh(b1,"#616161",btc)
        
        l1=tk.Label(fr1,text="_____________________________________________________________________________________________",bg=bgc,fg=bgc)
        l1.grid(row=1,column=0)
        
        b2=tk.Button(fr1,text="HWID Activation (For Windows 10, Lifetime)",command=lambda:(root.destroy(),run("hwidact")),borderwidth=0,bg=btc,fg="white",relief=RAISED)
        b2.grid(row=2,column=0)
        coh(b2,"#616161",btc)

        l2=tk.Label(fr1,text="_______________________________",bg=bgc,fg=bgc)
        l2.grid(row=3,column=0)

        b3=tk.Button(fr1,text="KMS38 Activation (For Windows 10, till 2038)",command=lambda:(root.destroy(),run("kms38act")),borderwidth=0,bg=btc,fg="white",relief=RAISED)
        b3.grid(row=4,column=0)
        coh(b3,"#616161",btc)

        l3=tk.Label(fr1,text="_______________________________",bg=bgc,fg=bgc)
        l3.grid(row=5,column=0)

        b4=tk.Button(fr1,text="Online KMS Activation (For All Windows and Office VL versions)",command=lambda:(onlkmsact()),borderwidth=0,bg=btc,fg="white",relief=RAISED)
        b4.grid(row=6,column=0)
        coh(b4,"#616161",btc)

        l4=tk.Label(fr1,text="_______________________________",bg=bgc,fg=bgc)
        l4.grid(row=7,column=0)

        b5=tk.Button(fr1,text="Extras",command=lambda:(extras()),borderwidth=0,bg=btc,fg="white",relief=RAISED)
        b5.grid(row=8,column=0)
        coh(b5,"#616161",btc)

def isodowner():
    fr3.destroy()
    fr4.grid(row=0,column=0)
    fr4.configure(bg=bgc)

    l1=tk.Label(fr4,text="_____________________________________________________________________________________________",bg=bgc,fg=bgc)
    l1.grid(row=0,column=0)

    global tv
    tv=ttk.Treeview(fr4,selectmode="browse",height=12)
    tv.grid(row=0,column=0,columnspan=2)
    tv["columns"] = ("1","2")
    tv.column("1")
    tv.column("2")
    tv.heading("1",text="ISO_ID")
    tv.heading("2",text="ISO_Name")

    p=open(rp("PyWinAct_Files/win.csv"),"r")  
    d=csv.reader(p)
    l=list(d)
    c=1
    for i in l:
        tv.insert("","end",values =(i[0],i[1]))
        c+=1
    e1=Entry(fr4,relief=RAISED)
    e1.grid(row=1,column=0,sticky="nsew")
    e1.insert("0","Enter ISO ID here(Delete this before typing)")
    print(e1.get())
    b1=Button(fr4,text="Download",bg=btc,fg=fgc,borderwidth=0,command=lambda:(wisodown("w",e1.get()),e1.delete(0,10)))
    b1.grid(row=1,column=1,sticky="nsew")
    coh(b1,"#616161",btc)
    b2t="""Click here if any of the
given links are not working"""
    b2=Button(fr4,text=b2t,bg=bgc,fg="#babab8",borderwidth=0,font=("TkDefaultFont",9,'underline'),command=lambda:(webbrowser.open("https://tb.rg-adguard.net/public.php")),relief=RAISED,cursor="dotbox")
    b2.grid(row=2,column=1,sticky="nsew")
    cohf(b2,fgc,"#babab8")

    l1=Label(fr4,text="ⓘ Enter ISO ID above and click Download",bg=bgc)  
    l1.grid(row=2,column=0,columnspan=2)

    b4=tk.Button(fr4,text="Back",command=lambda:(fr4.destroy(),mainmenu()),borderwidth=0,bg="red",fg="white")
    b4.grid(row=4,column=0,sticky="w")
    coh(b4,"#616161","red")
    
def wisodown(d,c):
    if d=="w":
        p=open(rp("PyWinAct_Files/win.csv"),"r")    
    elif d=="o":
        p=open(rp("PyWinAct_Files/office.csv"),"r")
    d=csv.reader(p)
    l=list(d)
    temp=[]
    for i in l:
        if i[0]==c:    
            webbrowser.open(i[2])
        temp.append(i[0])
    if str(c) not in temp:
        messagebox.showerror("Error","The ISO ID doesnt even exist dude!"+"/n"+"Try another one")

def oisodowner():
    fr3.destroy()
    fr4.grid(row=0,column=0)
    fr4.configure(bg=bgc)

    


    l1=tk.Label(fr4,text="_____________________________________________________________________________________________",bg=bgc,fg=bgc)
    l1.grid(row=0,column=0)

    global tv
    tv=ttk.Treeview(fr4,selectmode="browse",height=6)
    tv.grid(row=0,column=0,columnspan=2)
    tv["columns"] = ("1","2")
    tv.column("1")
    tv.column("2")
    tv.heading("1",text="ISO_ID")
    tv.heading("2",text="ISO_Name")

    p=open(rp("PyWinAct_Files/office.csv"),"r")  
    d=csv.reader(p)
    l=list(d)
    c=1
    for i in l:
        tv.insert("","end",values =(i[0],i[1]))
        c+=1
    e1=Entry(fr4,relief=RAISED)
    e1.grid(row=1,column=0,sticky="nsew")
    e1.insert("0","Enter ISO ID here(Delete this before typing)")
    print(e1.get())
    b1=Button(fr4,text="Download",bg=btc,fg=fgc,borderwidth=0,command=lambda:(wisodown("o",e1.get()),e1.delete(0,10)))
    b1.grid(row=1,column=1,sticky="nsew")
    coh(b1,"#616161",btc)
    b2t="""Click here if any of the 
given links are not working"""
    b2=Button(fr4,text=b2t,bg=bgc,fg="#babab8",borderwidth=0,font=("TkDefaultFont",9,'underline'),command=lambda:(webbrowser.open("https://tb.rg-adguard.net/public.php")),relief=RAISED,cursor="dotbox")
    b2.grid(row=2,column=1,sticky="nsew")
    cohf(b2,fgc,"#babab8")

    l1=Label(fr4,text="ⓘ Enter ISO ID above and click Download",bg=bgc)  
    l1.grid(row=2,column=0,columnspan=2)

    b4=tk.Button(fr4,text="Back",command=lambda:(fr4.destroy(),mainmenu()),borderwidth=0,bg="red",fg="white")
    b4.grid(row=4,column=0,sticky="w")
    coh(b4,"#616161","red")
 
def menubars():
    #Menubars Used
    global menubar
    menubar = tk.Menu(root,background=bgc,fg="white")
    #Extras Menu
    Extras = tk.Menu(menubar, tearoff=False,bg=bgc,fg="white")
    Extras.add_command(label="HWID Readme",command=lambda:(run("hwidr")),activebackground="#383842")
    Extras.add_command(label="KMS38 Readme",command=lambda:(run("kms38r")),activebackground="#383842")
    Extras.add_command(label="Online KMS Readme",command=lambda:(run("oldkmsr")),activebackground="#383842")
    Extras.add_separator()
    Extras.add_command(label="Extract OEM Readme",command=lambda:(run("extarctoemr")),activebackground="#383842")
    Extras.add_command(label="KMS38 Protection Readme",command=lambda:(run("kms38protr")),activebackground="#383842")
    menubar.add_cascade(label="ReadMe's", menu=Extras)
    #Help Menu
    helpmenu = tk.Menu(menubar, tearoff=False,bg=bgc,fg="white")
    helpmenu.add_command(label="Credits",activebackground="#383842",command=lambda:(run("credits")))
    helpmenu.add_command(label="About Creator",activebackground="#383842",command=githubme)

    menubar.add_cascade(label="Help",menu=helpmenu)

def githubme():
    webbrowser.open("https://github.com/geoseiden")

def run(a):
    if a=="check":
        subprocess.call(rp("PyWinAct_Files\Check_Activation.cmd"))
    elif a=="hwidact":
        subprocess.call(rp("PyWinAct_Files\HWID_Activation.cmd"))
    elif a=="kms38act":
        subprocess.call(rp("PyWinAct_Files\KMS38_Activation.cmd"))
    elif a=="onlkms":
        onlkmsact()
    elif a=="oldkmsact":
        subprocess.call(rp("PyWinAct_Files\Old-act\Activate.cmd"))
    elif a=="oldkmsren":
        subprocess.call(rp("PyWinAct_Files\Old-act\Renewal_Setup.cmd"))
    elif a=="oldkmsunin":
        subprocess.call(rp(r"PyWinAct_Files\Old-act\Uninstall.cmd"))
    elif a=="hwidr":
        d=str(rp(r"PyWinAct_Files\Readmefiles\hwid_readme.txt"))
        os.system(r"notepad.exe"+d)
    elif a=="kms38r":
        os.system("notepad.exe"+rp("PyWinAct_Files\Readmefiles\kms38_readme.txt"))
    elif a=="extractoemr":
        os.system("notepad.exe"+rp("PyWinAct_Files\Readmefiles\extractoem_readme.txt"))
    elif a=="kms38prot":
        os.system("notepad.exe"+rp("PyWinAct_Files\Readmefiles\kms38prot_readme.txt"))
    elif a=="oldkmsr":
        os.system("notepad.exe"+rp("PyWinAct_Files\Readmefiles\oldact_readme.txt"))
    elif a=="exOEM":
        subprocess.call(rp("PyWinAct_Files\Extras\Extract_OEM_Folder.cmd"))
    elif a=="win10install":
        subprocess.call(rp("PyWinAct_Files\Extras\OEMRET-Install_W10_Key.cmd"))
    elif a=="win10change":
        subprocess.call(rp("PyWinAct_Files\Extras\OEMRET-Change_W10_Edition.cmd"))
    elif a=="kms38prot":
        subprocess.call(rp("PyWinAct_Files\Extras\Protect_Unprotect-KMS38.cmd"))
    elif a=="abbodi":
        webbrowser.open("https://forums.mydigitallife.net/threads/74197/")
    elif a=="credits":
        os.system("notepad.exe"+rp("PyWinAct_Files\Readmefiles\Credits.txt"))
    elif a=="isodown":
        webbrowser.open("https://www.heidoc.net/php/Windows-ISO-Downloader.exe")

def onlkmsact():
    fr1.destroy()

    fr2.grid(row=0,column=0)
    fr2.configure(bg=bgc)

    l1=tk.Label(fr2,text="This will skip any permanent HWID and KMS38 Activation",bg=bgc,fg="white")
    l1.grid(row=1,column=0)
    
    l2=tk.Label(fr2,text='''KMS activates for 180 Days.Click on Activation Renewal for automatic renewal.
                          (For core/ProWMC edition it is 30/45 Days)''',bg=bgc,fg="white")
    l2.grid(row=2,column=0)
    
    l3=tk.Label(fr2,text="_____________________________________________________________________________________________",bg=bgc,fg=bgc)
    l3.grid(row=3,column=0)

    b1=tk.Button(fr2,text="Activate - Windows / Server / Office",command=lambda:run("oldkmsact"),borderwidth=0,bg=btc,fg="white")
    b1.grid(row=4,column=0)
    coh(b1,"#616161",btc)
    
    l4=tk.Label(fr2,text="_______________________________",bg=bgc,fg=bgc)
    l4.grid(row=5,column=0)

    b2=tk.Button(fr2,text="Activation Renewal",command=lambda:run("oldkmsren"),borderwidth=0,bg=btc,fg="white")
    b2.grid(row=6,column=0)
    coh(b2,"#616161",btc)

    l5=tk.Label(fr2,text="_______________________________",bg=bgc,fg=bgc)
    l5.grid(row=7,column=0)

    b3=tk.Button(fr2,text="Complete Uninstall",command=lambda:run("oldkmsunin"),borderwidth=0,bg=btc,fg="white")
    b3.grid(row=8,column=0)
    coh(b3,"#616161",btc)

    l6=tk.Label(fr2,text="_______________________________",bg=bgc,fg=bgc)
    l6.grid(row=9,column=0)

    b4=tk.Button(fr2,text="Back",command=lambda:(fr2.destroy(),mainmenu()),borderwidth=0,bg="red",fg="white")
    b4.grid(row=10,column=0,sticky="w")
    coh(b4,"#616161","red")

def extras():
    fr1.destroy()

    fr3.grid(row=0,column=0)
    fr3.configure(bg=bgc)

    l1=tk.Label(fr3,text="_____________________________________________________________________________________________",bg=bgc,fg=bgc)
    l1.grid(row=1,column=0)

    l2=tk.Label(fr3,text="_______________________________",bg=bgc,fg=bgc)
    l2.grid(row=3,column=0)

    b2=tk.Button(fr3,text="Insert Windows 10 Key [OEMRET]",command=lambda:(run("win10install")),borderwidth=0,bg=btc,fg="white")
    b2.grid(row=4,column=0)
    coh(b2,"#616161",btc)

    l3=tk.Label(fr3,text="_______________________________",bg=bgc,fg=bgc)
    l3.grid(row=5,column=0)

    b3=tk.Button(fr3,text="Change Windows 10 Edition [OEMRET]",command=lambda:(run("win10change")),borderwidth=0,bg=btc,fg="white")
    b3.grid(row=6,column=0)
    coh(b3,"#616161",btc)

    l4=tk.Label(fr3,text="_______________________________",bg=bgc,fg=bgc)
    l4.grid(row=7,column=0)

    b4=tk.Button(fr3,text="Protect / Unprotect KMS38 Activation",command=lambda:(run("kms38prot")),borderwidth=0,bg=btc,fg="white")
    b4.grid(row=8,column=0)
    coh(b4,"#616161",btc)
    
    l7=tk.Label(fr3,text="_______________________________",bg=bgc,fg=bgc)
    l7.grid(row=11,column=0)

    b6=tk.Button(fr3,text="Windows ISO Downloader",command=isodowner,borderwidth=0,bg=btc,fg="white")
    b6.grid(row=12,column=0)
    coh(b6,"#616161",btc)
    
    l7=tk.Label(fr3,text="_______________________________",bg=bgc,fg=bgc)
    l7.grid(row=13,column=0)

    b7=tk.Button(fr3,text="Office ISO Downloader",command=oisodowner,borderwidth=0,bg=btc,fg="white")
    b7.grid(row=14,column=0)
    coh(b7,"#616161",btc)
    
    l8=tk.Label(fr3,text="_______________________________",bg=bgc,fg=bgc)
    l8.grid(row=15,column=0)

    b7=tk.Button(fr3,text="Back",command=lambda:(fr3.destroy(),mainmenu()),borderwidth=0,bg="red",fg="white")
    b7.grid(row=16,column=0,sticky="w")
    coh(b7,"#616161","red")
    
def coh(button, colorOnHover, colorOnLeave):

    button.bind("<Enter>", func=lambda e: button.config(
        background=colorOnHover))
  
    button.bind("<Leave>", func=lambda e: button.config(
        background=colorOnLeave))

def cohf(button, colorOnHover, colorOnLeave):

    button.bind("<Enter>", func=lambda e: button.config(
        fg=colorOnHover))
  
    button.bind("<Leave>", func=lambda e: button.config(
        fg=colorOnLeave))

mainmenu()