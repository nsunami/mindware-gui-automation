# automate mindware event file export

import pyautogui, time, os, re
pyautogui.PAUSE = .2
pyautogui.FAILSAFE = True

width, height = pyautogui.size()
# Delay in the automation process
# the program starts to run after 2 seconds of clicking
time.sleep(2)

## This is the folder with the .mwi file.
# Change it to reflect the directly
mwi_dir = "W:\OneDrive - CRHLab\Study Materials\EVv1 (Nami) - 2017 Spring\Data - Electronic and Recruitment\Mindware"

# List of all the files in the .mwi directory
files = os.listdir(mwi_dir)

# Initialize the list of .mwi files
mwi_lst = []
# Iterate the file names to get the list of .mwi files
for name in files:
      if name.endswith('.mwi'):
            mwi_lst.append(name)

# Identify already-existing .txt files
# Initialize the list of text files
txt_lst = []
# Initialize the list of participant ids
ids = []
for txt in files:
      if txt.endswith('.txt'):
            txt_lst.append(txt)
            ids.append(txt.partition('_event')[0])

# List of .mwi files that do not have corresponding .txt files
mwi_remain = []
for mwi in mwi_lst:
      if mwi.partition('.')[0] not in ids:
            mwi_remain.append(mwi)


# Debug mode switch
# If on = the program asks to continue after each .mwi file
debug = 0


# Start the GUI automation
for mwi in mwi_remain:
    # Go to File -> Open to open the dialog box
    pyautogui.hotkey('alt')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('down')
    pyautogui.hotkey('enter')
    # Wait for one sec
    time.sleep(1)
    # Write the filename into the box
    pyautogui.typewrite(mwi)
    pyautogui.hotkey('enter')
    # Wait for one sec
    time.sleep(1)
    # The channel mapping window opens by default.
    # Map the channels
    # Channel 1 = ECG
    pyautogui.hotkey('tab')
    pyautogui.hotkey('space')
    pyautogui.typewrite('ECG')
    pyautogui.hotkey('enter')
    # Channel 2 = resp
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('space')
    pyautogui.typewrite('resp')
    pyautogui.hotkey('enter')
    # Channel 3 = Zo
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('space')
    pyautogui.typewrite('zo')
    pyautogui.hotkey('enter')
    # Channel 4 = dZ
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('space')
    pyautogui.typewrite('dz')
    pyautogui.hotkey('enter')
    # Go to "OK" button (shift + tab x 5 times)
    pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('enter')
    # Back to the main window
    # Wait for one sec
    time.sleep(1)
    # Go to File --> Export Events
    pyautogui.hotkey('alt')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('down')
    pyautogui.hotkey('down')
    pyautogui.hotkey('down')
    pyautogui.hotkey('down')
    pyautogui.hotkey('enter')
    time.sleep(1)
    # Save Eport Events Window
    # Filename is auto-filld by default
    pyautogui.hotkey('enter')
    time.sleep(1)
    # If the file exists, "File Already Exists Error appears"
    # YES
    pyautogui.hotkey('enter')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('enter')
    time.sleep(1)
    # If debug mode is on, wait for the response.
    if debug == 1:
        proceed = pyautogui.confirm("Do you want to continue?")
          if proceed == 'OK':
                continue
          if proceed == None:
                break
