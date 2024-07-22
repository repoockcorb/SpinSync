import subprocess
import os

cur_dir = os.getcwd()
cur_dir = cur_dir.replace("\\", "/")
print(cur_dir)

pyinstaller_command = [
    "pyinstaller",
    "spinsync.py",
    "--name", "SpinSync",
    # "--onedir",
    "--onefile",
    "--windowed",
    "--icon="+cur_dir+"/images/spinsync_icon.ico",
    "--paths", cur_dir+"/.venv/Lib/site-packages",
    "--add-data", cur_dir+"/images/spinsync_icon.ico;./images",
    "--add-data", cur_dir+"/images/SpinSync_logo.png;./images",
    # "--add-data", cur_dir+"/images/windlogix_white.png;./images"
]

subprocess.call(pyinstaller_command)