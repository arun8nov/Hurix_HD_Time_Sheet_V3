import os
import re
import subprocess

with open('full_compiled.html', 'r', encoding='utf-8') as f:
    html = f.read()

trap = """
<script>
window.__captured_errors = [];
window.onerror = function(msg, url, line, col, err) {
  window.__captured_errors.push({msg: msg, line: line, col: col, err: String(err)});
  var errDiv = document.createElement('div');
  errDiv.id = 'CAPTURED_DOM_ERROR';
  errDiv.style.color = 'red';
  errDiv.textContent = 'CAPTURED_ERROR: ' + msg + ' at line ' + line + ':' + col;
  document.body.appendChild(errDiv);
};
document.addEventListener('DOMContentLoaded', function() {
  var okDiv = document.createElement('div');
  okDiv.id = 'DOM_LOADED_OK';
  okDiv.textContent = 'DOM_AND_SCRIPTS_LOADED_SUCCESSFULLY';
  document.body.appendChild(okDiv);
});
</script>
"""

html_with_trap = html.replace('<head>', '<head>' + trap)
with open('scratch/test_trap.html', 'w', encoding='utf-8') as f:
    f.write(html_with_trap)

trap_path = os.path.abspath('scratch/test_trap.html').replace('\\', '/')
file_url = f'file:///{trap_path}'

edge_path = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
cmd = [
    edge_path,
    '--headless',
    '--disable-gpu',
    '--dump-dom',
    file_url
]

res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
output = res.stdout

if 'CAPTURED_DOM_ERROR' in output:
    print("FAILED: Errors captured in DOM:")
    for line in output.splitlines():
        if 'CAPTURED_ERROR:' in line:
            print("  ", line)
elif 'DOM_LOADED_OK' in output:
    print("SUCCESS: DOM and all JavaScript scripts loaded and executed without any runtime errors!")
else:
    print("Notice: DOM dump output length:", len(output))

