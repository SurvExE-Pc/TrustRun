import sys
import os
import subprocess
file = str(sys.argv[1])[11:len(str(sys.argv[1]))-1]
subprocess.call("wscript.exe C:\SystemRun\RunAs.vbs "+file)
