# automate mindware event file export

import pyautogui, time, os, re
pyautogui.PAUSE = .2
pyautogui.FAILSAFE = True

width, height = pyautogui.size()
# Delay in the automation process
time.sleep(2)

## This is the folder with the .mwi file.
# Change it to adjust
mwi_dir = "W:\OneDrive - CRHLab\Study Materials\EVv1 (Nami) - 2017 Spring\Data - Electronic and Recruitment\Mindware"

files = os.listdir(mwi_dir)

mwi_lst = []
for name in files:
      if name.endswith('.mwi'):
            mwi_lst.append(name)

txt_lst = []
ids = []
for txt in files:
      if txt.endswith('.txt'):
            txt_lst.append(txt)
            ids.append(txt.partition('_event')[0])

mwi_remain = []
for mwi in mwi_lst:
      if mwi.partition('.')[0] not in ids:
            mwi_remain.append(mwi)


debug = 0
for mwi in mwi_remain:
      pyautogui.hotkey('alt')
      pyautogui.hotkey('enter')
      pyautogui.hotkey('down')
      pyautogui.hotkey('enter')
      time.sleep(1)
      if pyautogui.locateCenterOnScreen('open.png') != None:
            quit()
      pyautogui.typewrite(mwi)
      pyautogui.hotkey('enter')
      time.sleep(1)
      # the channel mapping window
      if pyautogui.locateCenterOnScreen('wrong.png') != None:
            for i in range(17):
                  pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')
            continue
      else:
            pyautogui.hotkey('tab')
            pyautogui.hotkey('space')
            pyautogui.typewrite('ECG')
            pyautogui.hotkey('enter')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('space')
            pyautogui.typewrite('resp')
            pyautogui.hotkey('enter')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('space')
            pyautogui.typewrite('zo')
            pyautogui.hotkey('enter')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('space')
            pyautogui.typewrite('dz')
            pyautogui.hotkey('enter')
            pyautogui.hotkey('shift', 'tab')
            pyautogui.hotkey('shift', 'tab')
            pyautogui.hotkey('shift', 'tab')
            pyautogui.hotkey('shift', 'tab')
            pyautogui.hotkey('shift', 'tab')
            pyautogui.hotkey('enter')
            ## going back to the main window
            time.sleep(1)
            pyautogui.hotkey('alt')
            pyautogui.hotkey('enter')
            pyautogui.hotkey('down')
            pyautogui.hotkey('down')
            pyautogui.hotkey('down')
            pyautogui.hotkey('down')
            pyautogui.hotkey('enter')
            time.sleep(1)
            pyautogui.hotkey('enter')
            time.sleep(1)
            pyautogui.hotkey('enter')
            pyautogui.hotkey('enter')
            pyautogui.hotkey('enter')
            time.sleep(1)
            if debug == 1:
                  proceed = input('continue? y/n ')
                  if proceed == 'y':
                        continue
                  if proceed == 'n':
                        break




# pyautogui.keyDown('alt')
# pyautogui.keyUp('alt')

# file = pyautogui.locateOnScreen('file.png')
# pyautogui.click(file[0], file[1])
