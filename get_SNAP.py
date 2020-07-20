from bs4 import BeautifulSoup
import urllib.request, urllib.error, urllib.parse
from get_GatewayID import getGatewayID
##get dbP/N
def getSNAP(atp):
	rackurl = "http://scmdb/py/rms/gateway/" + str(getGatewayID(atp))
	print(rackurl)
	soup = BeautifulSoup(urllib.request.urlopen(rackurl), 'html.parser')
	SNAP = soup.find_all("td", string="SNAP")
	date = SNAP[len(SNAP)-1].next_sibling.next_sibling.string
	label = ""
	for l in SNAP[len(SNAP)-1].next_sibling.next_sibling.next_sibling.next_sibling.stripped_strings:
		label = label + l
	link_s = "http://scmdb" + SNAP[len(SNAP)-1].next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.next_sibling.a.get("href")
	return label, link_s, date


# print(getSNAP(1157774))
