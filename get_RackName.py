import requests
import json

def getRackName(atp):
	parameters = {"crid": atp}
	# Make a get request with the parameters.
	response = requests.get("http://scmdb/py/api/synergy/cr_data", params=parameters)
	# Print the content of the response (the data the server returned)
	rackname = json.loads(response.content)
	return rackname['data']['rack']

# print(getRackName(1157774))