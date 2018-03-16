from Tkinter import *

import re, traceback

from time import sleep

import tkMessageBox
import urllib2
import sys
import os
from bs4 import BeautifulSoup
import platform
if 'Win' in platform.system():
##    import win32com.client
    from win32com.client import Dispatch, constants
    import win32gui, win32con, win32com.client

##set window focus
class cWindow:
    def __init__(self):
        self._hwnd = None
        self.shell = win32com.client.Dispatch("WScript.Shell")

    def BringToTop(self):
        win32gui.BringWindowToTop(self._hwnd)

    def SetAsForegroundWindow(self):
        self.shell.SendKeys('%')
        win32gui.SetForegroundWindow(self._hwnd)

    def Maximize(self):
        win32gui.ShowWindow(self._hwnd, win32con.SW_MAXIMIZE)

    def setActWin(self):
        win32gui.SetActiveWindow(self._hwnd)

    def _window_enum_callback(self, hwnd, wildcard):
        '''Pass to win32gui.EnumWindows() to check all the opened windows'''
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._hwnd = hwnd

    def find_window_wildcard(self, wildcard):
        self._hwnd = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def kill_task_manager(self):
        wildcard = 'Gestionnaire des t.+ches de Windows'
        self.find_window_wildcard(wildcard)
        if self._hwnd:
            win32gui.PostMessage(self._hwnd, win32con.WM_CLOSE, 0, 0)
            sleep(0.5)

def wnd_d():
    sleep(5)
    try:
        wildcard = ".*DATABASE MANIFEST REQUEST for.*"
        cW = cWindow()
        cW.kill_task_manager()
        cW.find_window_wildcard(wildcard)
        cW.BringToTop()
##        cW.Maximize()
        cW.SetAsForegroundWindow()


    except:
        f = open("log.txt", "w")
        f.write(traceback.format_exc())
        print(traceback.format_exc())

def wnd_mani():
    sleep(5)
    try:
        wildcard1 = ".*MEGA-MANIFEST REQUEST for.*"
        cW1 = cWindow()
        cW1.kill_task_manager()
        cW1.find_window_wildcard(wildcard1)
        cW1.BringToTop()
##        cW1.Maximize()
        cW1.SetAsForegroundWindow()


    except:
        f = open("log.txt", "w")
        f.write(traceback.format_exc())
        print(traceback.format_exc())


##get ATP#
def getATP(buildinfourl):
    soup = BeautifulSoup(buildinfourl)
    panel_body = soup.find("div",{"class":"panel-body"})
    table_content = panel_body.form.find_all("div",{"class":"form-group"})
    last_div = None
    for last_div in table_content:pass
    return last_div.div.table.tr.next_sibling.next_sibling.find("td",{"class":"atp-title"}).get("data-atp")

##get SIT#
def getSIT(buildinfourl):
    soup = BeautifulSoup(buildinfourl)
    panel_body = soup.find("div",{"class":"panel-body"})
    table_content = panel_body.form.find_all("div",{"class":"form-group"})
    last_div = None
    i=0
    s=""
    for last_div in table_content:
        i=i+1
        if i>5:break
##    if last_div.div.p.a==None:
##        return "None"
    for string in last_div.div.p.stripped_strings:
        s=s+string
    if not s: s = "None"
    return s

##get P/N
def getPN(manifesturl):
    soup = BeautifulSoup(manifesturl)
    email_release = soup.find(title="Release Email Draft")
    email_release_p = email_release.find_all("p")
    s=""
    for p in email_release_p:
        for string in p.stripped_strings:
            s=s+string+"<br>"
    return s

##get Distributed info
def getDist(buildmemo):
    soup = BeautifulSoup(buildmemo)
    panel_body = soup.find("div",{"class":"panel-body"})
    dist_sec = panel_body.find("div",{"class":"form-group"})
    return dist_sec.p.string

##get list
def getList(buildlist):
    soup = BeautifulSoup(buildlist)
    panel_body = soup.find("div",{"class":"panel-body"})
    table_content = panel_body.find("div",{"class":"table-responsive"})
    item_to_mf = table_content.div.table.tbody.find_all("tr")
    name = []
    link = []
    for tr in item_to_mf:
        td = tr.find_all("td")
        if len(td)<3:continue
        try:
            if "No" in td[7].string:continue
        except:pass
        for string in td[3].span.a.stripped_strings:
            name.append(string)
        link.append(td[3].span.a.get("href"))
        
    title = name[0]
    name = name[1:]
    link = link[1:]
    name1=[]
    link1=[]
    d_name1=[]
    d_link1=[]
    for x in range(len(name)):
        if "44" in name[x]:
            if "S/W" in name[x]:continue
            if "IFE DB" in name[x]:
                d_name1.append(name[x][3:])
                d_link1.append("http://scmdb"+link[x])
                continue
            if "GCS DB" in name[x]:
                d_name1.append(name[x][3:])
                d_link1.append("http://scmdb"+link[x])
                continue
            if "EXW DB" in name[x]:
                d_name1.append(name[x][3:])
                d_link1.append("http://scmdb"+link[x])
                continue
            name1.append(name[x][3:])
            link1.append("http://scmdb"+link[x]) 
    return name1,link1,d_name1,d_link1,title  

##generate url
def getURL(buildno):
    buildinfo = "http://scmdb/py/scmbuild/"+buildno
    buildlist = "http://scmdb/py/scmbuild/"+buildno+"/show_details_list"
    buildinfo_s = "http://scmdb/py/scmbuild/"+buildno+"/show_information"
    buildmemo = "http://scmdb/py/scmbuild/"+buildno+"/show_memo_information"
    return buildinfo, buildlist, buildinfo_s, buildmemo

##testing 1184000b

##main
##if len(sys.argv)<2:
##    sys.exit("Need to attach a build number! for example:\npython maniFest9.py 1039003a")
##if len(sys.argv)>2:
##    sys.exit("Only one build number is allowed!")

def main(buildnumber):
    
    

    
    buildnumber = buildnumber.replace(' ','')


    
    L1.grid(row=3, column=1)
    master.update()
    print '\ngetting build '+ buildnumber +'....'
    


    
    bi,bl,bs,bm = getURL(buildnumber)
    buildlist = urllib2.urlopen(bl)##'buildlist.html'
    buildinfo_s = urllib2.urlopen(bs)
    buildmemo = urllib2.urlopen(bm)    

    ##bi,bl,bs = getURL(str(sys.argv[1]))
    ##buildlist = urllib2.urlopen(bl)##'buildlist.html'

    ##quote_page = 'http://docs.python-guide.org/en/latest/scenarios/scrape/'
    ##page = urllib2.urlopen(quote_page)
    L1.grid_forget()
    master.update()
    result = ""
    if getDist(buildmemo)=='No':
        result = tkMessageBox.askquestion("Distributed not checked",\
                                          "The build has not been distributed, continue manifest?",\
                                          icon='warning')
        if result == 'no':return
    
    L2.grid(row=3, column=1)
    master.update()    

    print 'getting manifest items...'
    name,link,d_name,d_link,title = getList(buildlist)
    print name, link ,d_name, d_link, title
    title = title[4:]
    ##open the manifestlink

    ##for z in link:

    ##    link[i] = urllib2.urlopen(link[i])
    L2.grid_forget()
    master.update()
    
    L3.grid(row=3, column=1)
    master.update()  
    print 'getting ATP#...'
    atp = getATP(buildinfo_s)
    buildinfo_s = urllib2.urlopen(bs)
    L3.grid_forget()
    master.update()
    
    L4.grid(row=3, column=1)
    master.update() 
    print 'getting SIT#...'
    sit = getSIT(buildinfo_s)
    L4.grid_forget()
    master.update()
    
    L5.grid(row=3, column=1)
    master.update() 
    print 'getting PN...'
    mani = []
    for x in link:
        tmp = urllib2.urlopen(x)
        mani.append(getPN(tmp))

    mani_d = []
    for y in d_link:
        tmp_d = urllib2.urlopen(y)
        mani_d.append(getPN(tmp_d))
    L5.grid_forget()
    master.update()
    
    L6.grid(row=3, column=1)
    master.update() 
    print 'composing email...'
    ##compose e-mail
    ##components text composing

    email_html = ""
    if name:
##        print "not name"
        email_html = email_html + "Dear SCM,<br><br>Could you please manifest following components below for "+\
                     "<a href="+bi+">build "+buildnumber+"</a>"+"<br><br>"
        i=0
        for x in name:
            email_html = email_html+"<a href="+link[i]+">"+x+"</a>"+"<br><br>"
            i=i+1
        email_html = email_html + "<br><br>ATP number is "+atp+"<br>"+\
                     "SIT number is "+sit+"<br />"+\
                     "The rack scan is <font color='red'>MISSING RACK SCAN HERE!!!DO NOT SEND OUT!!!</font>"+"<br><br><br>"
        email_html = email_html + "Part number information below.<br><br>"
        i=0
        for y in mani:
            email_html = email_html + name[i]+ "<br>"
            email_html = email_html + y + "<br><br>"
            i=i+1
        email_html = email_html + "<br>Thanks,<br>"

    ##database text composing
    email_d = ""
    if d_name:
##        print "not d_name"
        email_d = email_d + "Dear SCM,<br><br>Could you please manifest following database(s) for "+\
                  "<a href="+bi+">build "+buildnumber+"</a>"+"<br><br>"
        j=0
        for y in d_name:
            email_d = email_d +"<a href="+d_link[j]+">"+y+"</a>"  +", "+ \
                      "<br>ECSRR: <font color='red'>MISSING ECSRR HERE!!!DO NOT SEND OUT!!!</font><br><br>"+"<br>"
            j=j+1
        email_d = email_d + "ATP number is "+atp+"<br><br><br>"+"Part number information below.<br><br>"
        p=0
        for z in mani_d:
            email_d = email_d + d_name[p]+"<br>"
            email_d = email_d + z + "<br><br>"
            p=p+1
        email_d = email_d + "<br>Thanks,<br>" 
    
##    print 'successful!'
    title = title + " (ATP#"+ atp + " build#" + buildnumber +")"
    if result:
        title = title+" (Distributed: Not checked)"
    ##print email
    if 'Win' in platform.system():
        const=win32com.client.constants
        olMailItem = 0x0
        
        if email_html:
            
            obj = win32com.client.Dispatch("Outlook.Application")
            newMail = obj.CreateItem(olMailItem)
            newMail.Subject = "MEGA-MANIFEST REQUEST for "+title
            # newMail.Body = "I AM\nTHE BODY MESSAGE!"

            newMail.HTMLBody = email_html
            newMail.To = "socal.scm.ManifestRequest@panasonic.aero"
            newMail.display()
            wnd_mani()

##        const=win32com.client.constants
##        olMailItem = 0x0
        if email_d:
            
            obj = win32com.client.Dispatch("Outlook.Application")
            newMail = obj.CreateItem(olMailItem)
            newMail.Subject = "DATABASE MANIFEST REQUEST for "+title

            newMail.HTMLBody = email_d
            newMail.To = "socal.scm.ManifestRequest@panasonic.aero"
            ##attachment1 = r"C:\Temp\example.pdf"
            ##newMail.Attachments.Add(Source=attachment1)
            newMail.display()
            wnd_d()
        
    else:
        if email_html:
            cmd = """osascript -e 'tell application "Microsoft Outlook"' -e 'set newMessage to make new outgoing message with properties {subject:"MEGA-MANIFEST REQUEST for %s", content:"%s"}' -e 'make new recipient at newMessage with properties {email address:{address:"socal.scm.ManifestRequest@panasonic.aero"}}' -e 'open newMessage' -e 'end tell'""" %(title,email_html)
            os.system(cmd)
        if email_d:
            cmd1 = """osascript -e 'tell application "Microsoft Outlook"' -e 'set newMessage to make new outgoing message with properties {subject:"DATABASE MANIFEST REQUEST for %s", content:"%s"}' -e 'make new recipient at newMessage with properties {email address:{address:"socal.scm.ManifestRequest@panasonic.aero"}}' -e 'open newMessage' -e 'end tell'""" %(title,email_d)
            os.system(cmd1)
    L6.grid_forget()
    master.update()
def main_gui():
    try:
##        L1 = Label(master, text="working!")
##        L1.grid(row=3, column=1)
        

        master.update_idletasks()
        
        main(e1.get())
##        L1.destroy()
    except Exception as e:
        
##        L1.destroy()
        L1.grid_forget()
        L2.grid_forget()
        L3.grid_forget()
        L4.grid_forget()
        L5.grid_forget()
        L6.grid_forget()
        tkMessageBox.showinfo("Error", str(e))
        master.update()
def short_key(event):
    main_gui()


master = Tk()
master.title("MyManifest")
master.iconbitmap('if__m_2560433.ico')
Label(master, text="Build#").grid(row=0)

L1 = Label(master, text="getting build...")
L2 = Label(master, text="getting manifest items...")
L3 = Label(master, text="getting ATP#...")
L4 = Label(master, text="getting SIT#...")
L5 = Label(master, text="getting PN...")
L6 = Label(master, text="composing email...")

e1 = Entry(master)
e1.focus_set()


e1.grid(row=0, column=1)


##b1 = Button(master, text='Quit', command=master.quit)
##b1.grid(row=3, column=0, sticky=W, pady=4)
b2 = Button(master, text='Go', command=main_gui)
b2.grid(row=3, column=0, sticky=W, pady=4)

master.bind('<Return>',short_key)

mainloop( )

