import urllib.request, base64, json, urllib.parse

CLIENT_ID = '9bb0c2b3650f492f829320ac1c959171'
CLIENT_SECRET = '064eecd9c4024184aaff9cb02b76dd1c'

data = urllib.parse.urlencode({'grant_type': 'client_credentials'}).encode()
creds = base64.b64encode(f'{CLIENT_ID}:{CLIENT_SECRET}'.encode()).decode()
req = urllib.request.Request(
    'https://accounts.spotify.com/api/token',
    data=data,
    headers={'Authorization': f'Basic {creds}', 'Content-Type': 'application/x-www-form-urlencoded'}
)
token = json.loads(urllib.request.urlopen(req).read())['access_token']

req2 = urllib.request.Request(
    'https://api.spotify.com/v1/playlists/6cnvZnpG43x2ZMF3fFcsIo/tracks',
    headers={'Authorization': f'Bearer {token}'}
)
data2 = json.loads(urllib.request.urlopen(req2).read())
for t in data2['items']:
    print(t['track']['id'], '|', t['track']['name'])