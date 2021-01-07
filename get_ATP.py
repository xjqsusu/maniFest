from bs4 import BeautifulSoup
import urllib.request
##get ATP#
def getATP(buildinfourl):
    soup = BeautifulSoup(urllib.request.urlopen(buildinfourl), 'html.parser')
    # panel_body = soup.find("div",{"class":"panel-body"})
    # table_content = panel_body.form.find_all("div",{"class":"form-group"})
    # last_div = None
    # for last_div in table_content:pass
    # return last_div.div.table.tr.next_sibling.next_sibling.find("td",{"class":"atp-title"}).get("data-atp")
    try:
        s = soup.find("td",{"class":"atp-title"}).get("data-atp")
    except Exception:
        s = "Unknown"
    return s

# print(getATP("http://scmdb/py/scmbuild/608008a/show_information"))