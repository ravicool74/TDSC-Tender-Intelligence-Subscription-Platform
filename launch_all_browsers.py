import os
import subprocess
import shutil

url = "http://localhost:8000"

candidates = [
    ("Google Chrome", [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")
    ], "chrome.exe"),
    ("Microsoft Edge", [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    ], "msedge.exe"),
    ("Mozilla Firefox", [
        r"C:\Program Files\Mozilla Firefox\firefox.exe",
        r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
    ], "firefox.exe"),
    ("Brave", [
        r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\BraveSoftware\Brave-Browser\Application\brave.exe")
    ], "brave.exe"),
    ("Opera", [
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Opera\opera.exe"),
        r"C:\Program Files\Opera\opera.exe"
    ], "opera.exe")
]

opened = []
for name, paths, exe in candidates:
    target = None
    for p in paths:
        if os.path.exists(p):
            target = p
            break
    if not target:
        found_in_path = shutil.which(exe)
        if found_in_path:
            target = found_in_path

    if target:
        try:
            subprocess.Popen([target, url])
            opened.append(name)
            print(f"[SUCCESS] Launched {name}: {target}")
        except Exception as e:
            print(f"[ERROR] Failed {name}: {e}")

if not opened:
    import webbrowser
    webbrowser.open(url)
    print("[SUCCESS] Launched default browser")

print(f"\nTotal browsers launched: {len(opened)} ({', '.join(opened)})")
