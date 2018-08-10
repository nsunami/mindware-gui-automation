# Automate MindWare ASCII File Export


import time, os, re
import pyautogui

pyautogui.PAUSE = .8
pyautogui.FAILSAFE = True

# This script will get a list of .mwx file from a folder
#current_path = 'W:\GitHub\mindware-gui-automation'
#os.chdir(current_path)



# define the 
width, height = pyautogui.size()
# Delay in the automation process
# the program starts to run after 2 seconds of clicking
time.sleep(2)

## This is the folder with the .mwi file.
# Change it to reflect the directly
mwi_dir = "W:\OneDrive - CRHLab\Study Materials\EVv1 (Nami) - 2017 Spring\Data - Electronic and Recruitment\Mindware"
# the destination folder
txt_dir = mwi_dir + '\mwx2txt'

if os.path.exists(txt_dir) == False:
    os.mkdir(txt_dir)

# List of all the files in the .mwi directory
files = os.listdir(mwi_dir)
# Initialize the list of .mwi files
mwi_lst = []
# Iterate the file names to get the list of .mwi files
for name in files:
      if name.endswith('.mwi'):
            mwi_lst.append(name)

# List of txt files in the output directory
files_finished = [w.replace('.txt', '') for w in os.listdir(txt_dir)]

# Get the list of files to be processed.
files_tbp = set(mwi_lst) - set(files_finished)


# Debug mode switch
# If on = the program asks to continue after each .mwi file
debug = 0

# Start the GUI automation
for mwi in files_tbp:
    # Show a dialog box to start
    startup = pyautogui.confirm('Make sure to activate the BioLab window after clickng OK. You have 2 seconds to do so. ')
    if startup == 'OK':
        continue
    else:
        pyautogui.alert('The process was stopped by user.')
        break

    # Go to File -> Open to open the dialog box
    pyautogui.hotkey('alt')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('down')
    pyautogui.hotkey('enter')
    # Wait for 2 sec
    time.sleep(1)
    # Write the filename into the box
    pyautogui.typewrite(mwi)
    time.sleep(1)
    pyautogui.hotkey('enter')
    
    # Wait for 2 sec
    time.sleep(1)
    # Press View Button
    view_button = None
    while view_button == None:
        view_button = pyautogui.locateCenterOnScreen('view.png')
    pyautogui.click(x=view_button[0], y=view_button[1], button='left')
    
    # 'View' Screen
    # wait for it to load
    time.sleep(1)
    sat_button = None
    while sat_button == None:
        sat_button = pyautogui.locateCenterOnScreen('save_all_text.png')
    pyautogui.click(x=sat_button[0], y=sat_button[1], button='left')
    

    # 'SAVE AS' Window
    #wait for it to load
    time.sleep(3)
    pyautogui.typewrite(txt_dir + '\\' + mwi + '.txt')
    time.sleep(.5)
    # press enter
    pyautogui.hotkey('enter')

    # It takes a while to export the file
    writing = 1
    while writing != None:
        writing = pyautogui.locateCenterOnScreen('writing_txt.png')
        time.sleep(1)

    # Close the window
    playback_window = 1
    while playback_window != None:
        pyautogui.hotkey('ctrl', 'w')
        playback_window = pyautogui.locateCenterOnScreen('save_all_text.png')
        time.sleep(1)    
    
    # If debug mode is on, wait for the response.
    if debug == 1:
        proceed = pyautogui.confirm('Do you want to continue?')
        if proceed == 'OK':
            continue
        else:
            pyautogui.alert('The process was stopped by user.')
            break
        
pyautogui.alert('Done!')
