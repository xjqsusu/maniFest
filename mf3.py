
import urllib2
import sys
import os
from bs4 import BeautifulSoup
import platform
if 'Win' in platform.system():
    import win32com.client
    from win32com.client import Dispatch, constants


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
    if last_div.div.p.a==None:
        return "None"
    for string in last_div.div.p.a.stripped_strings:
        s=s+string
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
            name1.append(name[x][3:])
            link1.append("http://scmdb"+link[x]) 
    return name1,link1,d_name1,d_link1,title  

##generate url
def getURL(buildno):
    buildinfo = "http://scmdb/py/scmbuild/"+buildno
    buildlist = "http://scmdb/py/scmbuild/"+buildno+"/show_details_list"
    buildinfo_s = "http://scmdb/py/scmbuild/"+buildno+"/show_information"
    return buildinfo, buildlist, buildinfo_s

##testing 1184000b

##main
if len(sys.argv)<2:
    sys.exit("Need to attach a build number! for example:\npython maniFest9.py 1039003a")
if len(sys.argv)>2:
    sys.exit("Only one build number is allowed!")

def main(buildnumber):
    print 'opening build list...'
    
    bi,bl,bs = getURL(buildnumber)
    buildlist = urllib2.urlopen(bl)##'buildlist.html'
    buildinfo_s = urllib2.urlopen(bs)
        

    ##bi,bl,bs = getURL(str(sys.argv[1]))
    ##buildlist = urllib2.urlopen(bl)##'buildlist.html'

    ##quote_page = 'http://docs.python-guide.org/en/latest/scenarios/scrape/'
    ##page = urllib2.urlopen(quote_page)

    print 'getting manifest items...'
    name,link,d_name,d_link,title = getList(buildlist)
    ##print name, link
    title = title[4:]
    ##open the manifestlink

    ##for z in link:

    ##    link[i] = urllib2.urlopen(link[i])
    print 'getting ATP#...'
    atp = getATP(buildinfo_s)
    buildinfo_s = urllib2.urlopen(bs)
    print 'getting SIT#...'
    sit = getSIT(buildinfo_s)
    print 'getting PN...'
    mani = []
    for x in link:
        tmp = urllib2.urlopen(x)
        mani.append(getPN(tmp))

    mani_d = []
    for y in d_link:
        tmp_d = urllib2.urlopen(y)
        mani_d.append(getPN(tmp_d))
    print 'composing email...'
    ##compose e-mail
    f = open('manifest.txt','w')
    f.write('Dear SCM,\nCould you please manifest following components below for ')
    email = 'Dear SCM,\nCould you please manifest following components below for '
    email_html = "Dear, SCM,<br><br>Could you please manifest following components below for "
    for x in name:
        f.write(x+', ')
        email = email+x+', '
        email_html = email_html+x+", "
    f.write('\n\nATP number is '+atp+'\n'+'SIT number is '+sit+'\n'+'the rack scan is '+'\n\n\n')
    email = email + '\n\nATP number is '+atp+'\n'+'SIT number is '+sit+'\n'+'the rack scan is '+'\n\n\n'
    email_html = email_html + "<br><br>ATP number is "+atp+"<br>"+\
                 "SIT number is "+sit+"<br />"+\
                 "The rack scan is <font color='red'>MISSING RACK SCAN HERE!!!DO NOT SEND OUT!!!</font>"+"<br><br><br>"
    f.write('Part number information below.\n\n')
    email = email + 'Part number information below.\n\n'
    email_html = email_html + "Part number information below.<br><br>"
    i=0
    for y in mani:
        f.write(name[i]+'\n')
        email = email + name[i]+'\n'
        email_html = email_html + name[i]+ "<br>"
        f.write(y+'\n\n')
        email = email + y+'\n\n'
        email_html = email_html + y + "<br><br>"
        i=i+1
    f.write('\nThanks,\n')
    email = email + '\nThanks,\n'
    email_html = email_html + "<br>Thanks,<br>"
    f.close()

    f = open('dbmanifest.txt','w')
    f.write('Dear SCM,\nCould you please manifest the following database(s) for ')
    email_d = "Dear SCM,<br><br>Could you please manifest following database(s) for "
    for y in d_name:
        f.write(y+', ')
        email_d = email_d + y +", "
    f.write('\n\nATP number is '+atp+'\n\n\n'+'Part number information below.\n\n')
    email_d = email_d + "<br><br>ATP number is "+atp+"<br><br><br>"+"Part number information below.<br><br>"
    p=0
    for z in mani_d:
        f.write('\nECSRR:\n\n'+d_name[p]+'\n')
        email_d = email_d + "<br>ECSRR is <font color='red'>MISSING ECSRR HERE!!!DO NOT SEND OUT!!!</font><br><br>"+d_name[p]+"<br>"
        f.write(z+'\n\n')
        email_d = email_d + z + "<br><br>"
        p=p+1
    f.write('\nThanks,\n')
    email_d = email_d + "<br>Thanks,<br>" 
    f.close()
    
    print 'successful!'

    ##print email
    if 'Win' in platform.system():
        const=win32com.client.constants
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = "Mega-manifest request for "+title
        # newMail.Body = "I AM\nTHE BODY MESSAGE!"

        newMail.HTMLBody = email_html
        newMail.To = "socal.scm.ManifestRequest@panasonic.aero"
        newMail.display()

##        const=win32com.client.constants
##        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = "Database manifest request for "+title

        newMail.HTMLBody = email_d
        newMail.To = "socal.scm.ManifestRequest@panasonic.aero"
        ##attachment1 = r"C:\Temp\example.pdf"
        ##newMail.Attachments.Add(Source=attachment1)
        newMail.display()
    else:
        cmd = """osascript -e 'tell application "Microsoft Outlook"' -e 'set newMessage to make new outgoing message with properties {subject:"Mega-manifest request for %s", content:"%s"}' -e 'make new recipient at newMessage with properties {email address:{address:"socal.scm.ManifestRequest@panasonic.aero"}}' -e 'open newMessage' -e 'end tell'""" %(title,email_html)
        cmd1 = """osascript -e 'tell application "Microsoft Outlook"' -e 'set newMessage to make new outgoing message with properties {subject:"Database manifest request for %s", content:"%s"}' -e 'make new recipient at newMessage with properties {email address:{address:"socal.scm.ManifestRequest@panasonic.aero"}}' -e 'open newMessage' -e 'end tell'""" %(title,email_d)
        os.system(cmd)
        os.system(cmd1)

try:
    main(str(sys.argv[1]))
except Exception as e:
    sys.exit(str(e))


    
##f = open('manifest.txt','w')
##f.write('Dear SCM,'+'\n')
##f.close()
##
##f = open('manifest.txt','a')
##f.write(soup.get_text().encode('utf-8'))
##f.close()
