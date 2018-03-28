# importing the requests library
import requests


# Find SKU of event
eventUrl = "https://api.vexdb.io/v1/get_events"
eventParams = {'date':'2018-04-25'}
events = requests.get(eventUrl,eventParams).json()['result']
eventsSize = len(events)

for x in xrange(eventsSize):
    print(events[x]['name'])

print('done')
"""
# api-endpoint
URL = "https://api.vexdb.io/v1/get_matches"


# defining a params dict for the parameters to be sent to the API
PARAMS = {'team':'2442B'}

# sending get request and saving the response as response object
r = requests.get(url = URL, params = PARAMS)

# extracting data in json format
data = r.json()

print(data)
print('status: ' + str(data['status']))
print('sku: ' + data['result'][0]['sku'])
"""
