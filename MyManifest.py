from tkinter import *
from get_database_PN import getDBPN
from get_ATP import getATP
from get_SNAP import getSNAP
import re, traceback

from time import sleep
from nameInOS import get_display_name
import tkinter.messagebox
import urllib.request, urllib.error, urllib.parse
import sys
import os
from bs4 import BeautifulSoup
import platform
if 'Win' in platform.system():
##    import win32com.client
    from win32com.client import Dispatch, constants
    import win32gui, win32con, win32com.client
from appscript import app, k


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
        # f = open("log.txt", "w")
        # f.write(traceback.format_exc())
        print((traceback.format_exc()))

def wnd_mani():
    sleep(5)
    try:
        wildcard1 = ".*MANIFEST REQUEST for.*"
        cW1 = cWindow()
        cW1.kill_task_manager()
        cW1.find_window_wildcard(wildcard1)
        cW1.BringToTop()
##        cW1.Maximize()
        cW1.SetAsForegroundWindow()


    except:
        # f = open("log.txt", "w")
        # f.write(traceback.format_exc())
        print((traceback.format_exc()))




##get SIT#
def getSIT(buildinfourl):
    soup = BeautifulSoup(urllib.request.urlopen(buildinfourl), 'html.parser')
#     panel_body = soup.find("div",{"class":"panel-body"})
#     table_content = panel_body.form.find_all("div",{"class":"form-group"})
#     last_div = None
#     i=0
#     s=""
#     for last_div in table_content:
#         i=i+1
#         if i>5:break
# #    if last_div.div.p.a==None:
# #        return "None"

#     for string in last_div.div.p.stripped_strings:
#         s=s+string
#     if not s: s="Unknown"
#     return s 
    s=""
    try:
        if soup.find("label",{"for":"core_sw"}) == None: 
            table_content = soup.find("label",{"for":"sit_build"})
        else:
            table_content = soup.find("label",{"for":"core_sw"})
        for string in table_content.next_sibling.next_sibling.p.stripped_strings:
            s=s+string
        if not s: s="Unknown"
    except Exception:
        s = "Unknown"
    return s

##get ETE build#
def getETE(buildinfourl):
    soup = BeautifulSoup(urllib.request.urlopen(buildinfourl), 'html.parser')
    s=""
    try:
        for string in soup.find("label",{"for":"ete_build"}).next_sibling.next_sibling.p.stripped_strings:
            s=s+string
        if not s: s="Unknown"
    except Exception:
        s="Unknown"
    return s

##get P/N
def getPN(manifesturl):
    soup = BeautifulSoup(manifesturl, 'html.parser')
    email_release = soup.find(title="Release Email Draft")
    email_release_p = email_release.find_all("p")
    s=""
    for p in email_release_p:
        for string in p.stripped_strings:
            s=s+string+"<br>"
    return s

##get Distributed info
def getDist(buildmemo):
    soup = BeautifulSoup(urllib.request.urlopen(buildmemo), 'html.parser')
    panel_body = soup.find("div",{"class":"panel-body"})
    dist_sec = panel_body.find("div",{"class":"form-group"})
    return dist_sec.p.string

##get list
def getList(buildlist, GCS_enable):
    soup = BeautifulSoup(urllib.request.urlopen(buildlist), 'html.parser')
    panel_body = soup.find("div",{"class":"panel-body"})
    table_content = panel_body.find("div",{"class":"table-responsive"})
    item_to_mf = table_content.div.table.tbody.find_all("tr")

    title_set =item_to_mf[0]
    title_line = title_set.find_all("td")
    title=""
    for str in title_line[3].stripped_strings:
        title=title+str

    if "SCI" not in title:
        title="the build"
    else:
        item_to_mf = item_to_mf[1:]
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
##    title = name[0]
##    name = name[1:]
##    link = link[1:]
    name1=[]
    link1=[]
    name2=[]
    link2=[]
    d_name1=[]
    d_link1=[]
    for x in range(len(name)):
        if name[x].startswith("44"):
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
            if "S/W" in name[x]:continue
            if "GEN3" in name[x]:
                if not GCS_enable: continue
                else: 
                    name2.append(name[x][3:])
                    link2.append("http://scmdb"+link[x])
                    continue
            name1.append(name[x][3:])
            link1.append("http://scmdb"+link[x]) 
    return name1,link1,d_name1,d_link1,title,name2,link2

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

def main(buildnumber, GCS_enable):
    
    

    
    buildnumber = buildnumber.replace(' ','')


    
    L1.grid(row=4, column=1)
    master.update()
    print(("\ngetting build "+ buildnumber +"...."))
    


    
    bi,buildlist,buildinfo_s,buildmemo = getURL(buildnumber)
       

    ##bi,bl,bs = getURL(str(sys.argv[1]))
    ##buildlist = urllib2.urlopen(bl)##'buildlist.html'

    ##quote_page = 'http://docs.python-guide.org/en/latest/scenarios/scrape/'
    ##page = urllib2.urlopen(quote_page)
    L1.grid_forget()
    master.update()
    result = ""
    if getDist(buildmemo)=='No':
        result = tkinter.messagebox.askquestion("Distributed not checked",\
                                          "The build has not been distributed, continue manifest?",\
                                          icon='warning')
        if result == 'no':return
    
    L2.grid(row=4, column=1)
    master.update()    

    print('getting manifest items...')
    name,link,d_name,d_link,title,name_gcs,link_gcs = getList(buildlist, GCS_enable)
    title = title[4:]
    ##open the manifestlink

    ##for z in link:

    ##    link[i] = urllib2.urlopen(link[i])
    L2.grid_forget()
    master.update()
    
    L3.grid(row=4, column=1)
    master.update()  
    print('getting ATP#...')
    atp = getATP(buildinfo_s)
    L3.grid_forget()
    master.update()
    
    L4.grid(row=4, column=1)
    master.update() 
    print('getting SIT#...')
    sit = getSIT(buildinfo_s)
    ete = getETE(buildinfo_s)
    L4.grid_forget()
    master.update()
    
    L5.grid(row=4, column=1)
    master.update() 
    print('getting PN...')
    mani = []
    for x in link:
        tmp = urllib.request.urlopen(x)
        try:
            mani.append(getPN(tmp))
        except Exception:
            pass

    mani_g = []
    for x in link_gcs:
        tmp = urllib.request.urlopen(x)
        try:
            mani_g.append(getPN(tmp))
        except Exception:
            pass

    mani_d = []
    ecsrr_No = []
    warning_state = []
    for y in d_link:
        tmp_d = urllib.request.urlopen(y)
        mani_d.append(getPN(tmp_d))
        try:
            ecsrr_No.append(getDBPN(y))
            warning_state.append("")
        except Exception:
            warning_state.append("<font color='red'>MISSING ECSRR HERE!!!DO NOT SEND OUT!!!</font>")
            ecsrr_No.append("")
            pass
    L5.grid_forget()
    master.update()
    
    L6.grid(row=4, column=1)
    master.update() 
    print('composing email...')
    ##compose e-mail
    ##components text composing

    # print(sit)

    try:
        [label_snap, link_snap, date] = getSNAP(atp)
        tkinter.messagebox.showinfo("Latest snapshot found", "Snapshot <" + label_snap + "> dated "+ date +" found")
    except Exception:
        label_snap = "<font color='red'>MISSING RACK SCAN HERE!!!DO NOT SEND OUT!!!</font>"
        link_snap = ""
        date = ""


    email_html = ""
    if name:
##        print "not name"
        email_html = email_html + "Hello SCM,<br><br>Could you please manifest following IFE component(s) for "+\
                     "<a href="+bi+">build "+buildnumber+"</a>"+"<br><br>"
        i=0
        for x in name:
            email_html = email_html+"<a href="+link[i]+">"+x+"</a>"+"<br><br>"
            i=i+1
        email_html = email_html + "<br><br>ATP number is "+atp+"<br>"+\
                     "SIT number is "+sit+"<br />"+\
                     "The rack scan is " + "<a href="+link_snap+">"+label_snap+"</a>" +"<br><br><br>"
        email_html = email_html + "Below is part number information.<br><br>"
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
        # print(d_link)
        # print(ecsrr_No)
        email_d = email_d + "Hello SCM,<br><br>Could you please manifest following database(s) for "+\
                  "<a href="+bi+">build "+buildnumber+"</a>"+"<br><br>"
        j=0
        for y in d_name:
            email_d = email_d +"<a href="+d_link[j]+">"+y+"</a>"  +", <br>"+ \
                      "ECSRR: " + warning_state[j] + ecsrr_No[j]+"<br><br><br>"
            j=j+1
        email_d = email_d + "ATP number is "+atp+"<br><br><br>"+"Below is part number information.<br><br>"
        p=0
        for z in mani_d:
            email_d = email_d + d_name[p]+"<br>"
            email_d = email_d + z + "<br><br>"
            p=p+1
        email_d = email_d + "<br>Thanks,<br>" 

    ##gcs email composing
    email_gcs = ""
    if name_gcs:
##        print "not name"
        email_gcs = email_gcs + "Hello SCM,<br><br>Could you please manifest following GCS component(s) for "+\
                     "<a href="+bi+">build "+buildnumber+"</a>"+"<br><br>"
        i=0
        for x in name_gcs:
            email_gcs = email_gcs+"<a href="+link_gcs[i]+">"+x+"</a>"+"<br><br>"
            i=i+1
        email_gcs = email_gcs + "<br><br>ATP number is "+atp+"<br>"+\
                     "ETE build is "+ete+"<br />"+\
                     "The rack scan is <font color='red'>MISSING RACK SCAN HERE!!!DO NOT SEND OUT!!!</font>"+"<br><br><br>"
        email_gcs = email_gcs + "Below is part number information.<br><br>"
        i=0
        for y in mani_g:
            email_gcs = email_gcs + name_gcs[i]+ "<br>"
            email_gcs = email_gcs + y + "<br><br>"
            i=i+1
        email_gcs = email_gcs + "<br>Thanks,<br>"
    
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
            newMail.Subject = "IFE s/w MANIFEST REQUEST for "+title
            # newMail.Body = "I AM\nTHE BODY MESSAGE!"

            newMail.HTMLBody = email_html + get_display_name()
            newMail.To = "socal.scm.ManifestRequest@panasonic.aero"
            newMail.display()
            wnd_mani()

##        const=win32com.client.constants
##        olMailItem = 0x0
        if email_d:
            
            obj = win32com.client.Dispatch("Outlook.Application")
            newMail = obj.CreateItem(olMailItem)
            newMail.Subject = "DATABASE MANIFEST REQUEST for "+title

            newMail.HTMLBody = email_d + get_display_name()
            newMail.To = "socal.scm.ManifestRequest@panasonic.aero"
            ##attachment1 = r"C:\Temp\example.pdf"
            ##newMail.Attachments.Add(Source=attachment1)
            newMail.display()
            wnd_d()

        if email_gcs:

            obj = win32com.client.Dispatch("Outlook.Application")
            newMail = obj.CreateItem(olMailItem)
            newMail.Subject = "GCS s/w MANIFEST REQUEST for "+title

            newMail.HTMLBody = email_gcs + get_display_name()
            newMail.To = "socal.scm.ManifestRequest@panasonic.aero"
            ##attachment1 = r"C:\Temp\example.pdf"
            ##newMail.Attachments.Add(Source=attachment1)
            newMail.display()
            wnd_d()
        
    else:
        # if email_html:
        #     cmd = """osascript -e 'tell application "Microsoft Outlook"' -e 'set newMessage to make new outgoing message with properties {subject:"MANIFEST REQUEST for %s", content:"%s"}' -e 'make new recipient at newMessage with properties {email address:{address:"socal.scm.ManifestRequest@panasonic.aero"}}' -e 'open newMessage' -e 'end tell'""" %(title,email_html)
        #     os.system(cmd)
        # if email_d:
        #     cmd1 = """osascript -e 'tell application "Microsoft Outlook"' -e 'set newMessage to make new outgoing message with properties {subject:"DATABASE MANIFEST REQUEST for %s", content:"%s"}' -e 'make new recipient at newMessage with properties {email address:{address:"socal.scm.ManifestRequest@panasonic.aero"}}' -e 'open newMessage' -e 'end tell'""" %(title,email_d)
        #     os.system(cmd1)
        title_sw = "IFE s/w MANIFEST REQUEST for "+title
        title_d = "DATABASE MANIFEST REQUEST for "+title
        # email_html = email_html + get_display_name()
        # email_d = email_d + get_display_name()
        outlook = app('Microsoft Outlook')
        msg = outlook.make(
            new=k.outgoing_message,
            with_properties={
                k.subject: title_sw,
                k.plain_text_content: email_html})
        msg.make(
            new=k.recipient,
            with_properties={
                k.email_address: {
                    k.address: 'socal.scm.ManifestRequest@panasonic.aero'}})
        msg.open()
        msg.activate()
        outlook = app('Microsoft Outlook')
        msg = outlook.make(
            new=k.outgoing_message,
            with_properties={
                k.subject: title_d,
                k.plain_text_content: email_d})
        msg.make(
            new=k.recipient,
            with_properties={
                k.email_address: {
                    k.address: 'socal.scm.ManifestRequest@panasonic.aero'}})
        msg.open()
        msg.activate()


    L6.grid_forget()
    master.update()
def main_gui():
    try:
##        L1 = Label(master, text="working!")
##        L1.grid(row=4, column=1)
        

        master.update_idletasks()
        
##        main(e1.get(), var1.get())
        main(e1.get(), 0)
##        L1.destroy()
    except Exception as e:
        
##        L1.destroy()
        L1.grid_forget()
        L2.grid_forget()
        L3.grid_forget()
        L4.grid_forget()
        L5.grid_forget()
        L6.grid_forget()
        tkinter.messagebox.showinfo("Error", str(e))
        master.update()
def short_key(event):
    main_gui()


master = Tk()
master.title("MyManifest")
master.iconbitmap('iconfinder__m_2560433.icns')
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
##b1.grid(row=4, column=0, sticky=W, pady=4)
b2 = Button(master, fg='black', text='Go', command=main_gui)
b2.grid(row=4, column=0, sticky=W, pady=4)

##var1 = IntVar()
##c1 = Checkbutton(master, text='include GCS s/w', variable=var1, onvalue=1, offvalue=0)
##c1.grid(row=3, column=1, sticky=W)

master.bind('<Return>',short_key)

mainloop( )

