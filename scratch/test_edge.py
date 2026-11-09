import subprocess
import time
import urllib.request
import json
import os

html_path = os.path.abspath('full_compiled.html').replace('\\', '/')
file_url = f'file:///{html_path}'

# Start msedge with remote debugging
edge_path = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
cmd = [
    edge_path,
    '--headless',
    '--disable-gpu',
    '--remote-debugging-port=9222',
    file_url
]

proc = subprocess.Popen(cmd)
time.sleep(2)

try:
    with urllib.request.urlopen('http://localhost:9222/json') as response:
        pages = json.loads(response.read().decode())
    print("Edge opened pages successfully:", len(pages))
    for p in pages:
        print("Page title:", p.get('title'), "URL:", p.get('url'))
finally:
    proc.terminate()
    proc.wait()

