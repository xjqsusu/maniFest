import requests
import json
from get_RackName import getRackName

def getGatewayID(atp):
	rackname = getRackName(atp)
	parameters = {"hostname": rackname}
	# Make a get request with the parameters.
	response = requests.get("http://scmdb//py/api/rms/gateway/hostname", params=parameters)
	# Print the content of the response (the data the server returned)
	gateway = json.loads(response.content)
	return gateway['gateway']['id']

# print(getGatewayID(1157774))