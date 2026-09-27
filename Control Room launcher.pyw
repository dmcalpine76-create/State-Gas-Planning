# Double-click to start the State Gas Control Room launcher (no window).
# It opens the Control Room in your browser, paired with this PC.
# To start it automatically at login: put a shortcut to this file in
# your Startup folder (Win+R, type shell:startup, press Enter).
import sys, runpy, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))
sys.argv = ["launcher.py", "--pair"]
runpy.run_path("launcher.py", run_name="__main__")
