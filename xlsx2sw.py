import os
import sys
import pandas as pd
from pySW import SW
import time
import psutil
import PySimpleGUI as sg

# Read Excel File
INPUT = sg.popup_get_file("Select xlsx file",  title="File selector")
#INPUT = sys.argv[1]
if INPUT is None:
    sys.exit()
DF1 = pd.read_excel(INPUT,skiprows=2)
# Drop row FLAG==NaN
DF1 = DF1.dropna(subset=["ACTIVATION"])

# SEED SLDPRT File Name
SEED = os.path.splitext(INPUT)[0]
SEED_BASENAME = os.path.basename(SEED)
SEED_DIR = os.path.dirname(SEED)
NAME = ""

# Start Solidworks
if "SLDWORKS.exe" in (p.name() for p in psutil.process_iter()) == False:
    print('starting SLDWORKS...')
    SW.startSW();
    time.sleep(10);
SW.connectToSW()

for i in range(len(DF1.index)):
    # Read SLDPRT
    SW.openPrt(SEED+".SLDPRT");
    # Get Global Variables
    GVAR = SW.getGlobalVars();
    GVAR_NAME = [];
    for key in GVAR:
        GVAR_NAME.append(key);
    # Modify Global Variables and NAME
    for j in range(len(DF1.columns)):
        for k in range(len(GVAR_NAME)):
            if DF1.columns[j] == GVAR_NAME[k]:
                SW.modifyGlobalVar(GVAR_NAME[k],DF1.iloc[i,j],'mm');
            # Get NEW_NAME
            elif DF1.columns[j] == "NAME":
                if pd.isna(DF1.iloc[i,j]):
                    NAME = "NaN";
                else:
                    NAME = DF1.iloc[i,j];
    SW.update();
    # Save Modified SLDPRT
    if NAME == "NaN":
        SAVE_DIR = SEED_DIR+"/"+str(i+1)
        SAVE_FILE0 = str(i+1)
        SAVE_FILE2 = str(i+1)+".SLDPRT"
        SAVE_FILE3 = str(i+1)+".STEP"
        SAVE_FILE4 = str(i+1)+".STL"
        SAVE_FILE = SAVE_DIR+"/"+SAVE_FILE2
    else:
        SAVE_DIR = SEED_DIR+"/"+NAME
        SAVE_FILE0 = NAME
        SAVE_FILE2 = NAME+".SLDPRT"
        SAVE_FILE3 = NAME+".STEP"
        SAVE_FILE4 = NAME+".STL"
        SAVE_FILE = SAVE_DIR+"/"+SAVE_FILE2
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)
    if os.path.exists(SAVE_FILE):
        os.remove(SAVE_FILE)
    SW.save(SAVE_DIR,SAVE_FILE0,'SLDPRT');
    print(SAVE_FILE2)
    SW.save(SAVE_DIR,SAVE_FILE0,'STEP');
    print(SAVE_FILE3)
    SW.save(SAVE_DIR,SAVE_FILE0,'STL');
    print(SAVE_FILE4)
    # Close Doc
    SW.closeDoc()

# Shutdown SLDWORKS
SW.shutSW();
time.sleep(1);

print("Finished...")
