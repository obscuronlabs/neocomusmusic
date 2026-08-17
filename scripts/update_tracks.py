import os
import json
import re
import urllib.request
import urllib.parse
import base64

CLIENT_ID = os.environ['SPOTIFY_CLIENT_ID']
CLIENT_SECRET = os.environ['SPOTIFY_CLIENT_SECRET']
PLAYLIST_ID = os.environ['SPOTIFY_PLAYLIST_ID']

# Token al
creds = base64.b64encode(f"{CLIENT_ID}:{CLIENT_SECRET}".encode()).decode()
req = urllib.request.Request(
    'https://accounts.spotify.com/api/token',
    data=b'grant_type=client_credentials',
    headers={'Authorization': f'Basic {creds}', 'Content-Type': 'application/x-www-form-urlencoded'}
)
token = json.loads(urllib.request.urlopen(req).read())['access_token']

# Playlist'teki şarkıları çek
req2 = urllib.request.Request(
    f'https://api.spotify.com/v1/playlists/{PLAYLIST_ID}/tracks?limit=50&fields=items(track(id,name))',
    headers={'Authorization': f'Bearer {token}'}
)
data = json.loads(urllib.request.urlopen(req2).read())

tracks = []
for item in data['items']:
    t = item.get('track')
    if t and t.get('id'):
        tracks.append({'id': t['id'], 'title': t['name'].upper()})

print(f"Fetched {len(tracks)} tracks from Spotify")

# Yeni JS tracks array oluştur
js_lines = ['const tracks = [']
for t in tracks:
    js_lines.append(f"  {{ id: '{t['id']}', title: '{t['title'].replace(chr(39), chr(92)+chr(39))}' }},")
js_lines.append('];')
new_tracks_js = '\n'.join(js_lines)

# index.html içindeki tracks array'i güncelle
with open('index.html', 'r', encoding='utf-8') as f:
    html = f.read()

html = re.sub(
    r'const tracks = \[.*?\];',
    new_tracks_js,
    html,
    flags=re.DOTALL
)

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("index.html updated successfully")
