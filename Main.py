#libary imports
from guizero import App, Combo, Text, CheckBox, ButtonGroup, PushButton, info
import glob
import os
import sys
#
os.system("mkdir TEMP")
#mdir="C:/Users/"+(os.getlogin())+"/Downloads/"
mdir="input/"
#if passed info eg:{python3 Main.py filelocation} skip GUI
if len(sys.argv) >= 2:
	if sys.argv[1] != "":
		fin=sys.argv[1]
		os.system("python3 dll/pdf2txt.py "+fin)
		exit()

def process():#Process command
	os.system("python3 dll/pdf2txt.py "+file_choice.value)
	info("Status", "done :)")

#GUI Stuff
app = App(title="Unca GUI", width=300, height=200, layout="grid")
file_choice = Combo(app, options=(glob.glob(mdir+"/*.pdf")), grid=[0,0], align="left")
file_caption = Text(app, text="What To Process?", grid=[0,1], align="left")
play_file = PushButton(app, command=process, text="Go", grid=[0,4], align="left")
app.display()
#os.system("rmdir TEMP") #not sure if I should add this yet
