# automate mindware event file export
import pyautogui as pag
import time, os, re
import numpy as np
pag.PAUSE = .3
pag.FAILSAFE = True

conf = .9

width, height = pag.size()
# Delay in the automation process
# the program starts to run after 2 seconds of clicking
time.sleep(2)

## This is the folder with the .mwi file.
# Change it to reflect the directly
#mwi_dir = "C:\\Users\\asdf\\OneDrive - CRHLab\\Study Materials\\EVv1 (Nami) - 2017 Spring\\Data - Electronic and Recruitment\Mindware"
mwi_dir = 'W:\OneDrive - CRHLab\Study Materials\EVv1 (Nami) - 2017 Spring\Data - Electronic and Recruitment\Mindware'
excel_dir = mwi_dir + "\Excel files\Auto"
pag.prompt(text = 'Confirm the folder containing the original mwi files',
 title = 'Confirm Folder', default = mwi_dir)


# List of all the files in the .mwi directory
files = os.listdir(mwi_dir)

# Initialize the list of .mwi files
mwi_lst = []
# Iterate the file names to get the list of .mwi files
for name in files:
      if name.endswith('.mwi'):
            mwi_lst.append(name)

# Identify already-existing .excel files
# List of txt files in the output directory
files_finished = [w.replace('.xlsx', '') for w in os.listdir(excel_dir)]

# Get the list of files to be processed.
files_tbp = set(mwi_lst) - set(files_finished)


# Get the list of participats with chest electrode measurements
chest_electrodes = np.loadtxt('chest_electrodes.csv', skiprows = 1, delimiter = ',')
ids = chest_electrodes[:,0].astype(int)

mwi_ids = []
for name in mwi_lst:
    #mwi_ids.append(re.sub('\_\d\.mwi', '', name))
    mwi_ids.append(re.search('[0-9]{2,6}', name).group(0))

mwi_ids_int = np.array(mwi_ids).astype(int)

missing = np.in1d(ids, mwi_ids_int)
missing_id = ids[np.invert(missing)]

# Debug mode switch
# If on = the program asks to continue after each .mwi file
debug = 1

# Show a dialog box to start
startup = pag.confirm('Make sure to activate the IMP window after clickng OK. You have 2 seconds to do so. ')
time.sleep(2)


# Grayscale match?
grayscale_bool = True

# Start the GUI automation
for mwi in files_tbp:

    # Look for the chest electrode distance based on pid
    pid = np.array(re.search('[0-9]{2,6}', mwi).group(0), dtype = int)
    index = np.where(chest_electrodes == pid)
    electrode_cm = chest_electrodes[index, 1][0]




    # Determine the safe click area to go back to the window
    clickhere = pag.locateCenterOnScreen('file_path.png', grayscale=grayscale_bool)
    if clickhere == None:
        pag.alert('Cannot see that IMP program is open. Process Terminated.')
        break
    pag.click(clickhere)

    # Make sure that the active tab is 'Events and Modes'
    pag.click(pag.locateCenterOnScreen('events_and_modes.png', grayscale=grayscale_bool), clicks = 2)

    # Go to File -> Open to open the dialog box
    pag.hotkey('alt')
    pag.hotkey('enter')
    pag.hotkey('down')
    pag.hotkey('enter')
    # Wait for one sec
    time.sleep(1)
    # Write the filename into the box
    pag.typewrite(mwi)
    pag.hotkey('enter')
    # Wait for one sec
    time.sleep(1)
    # The channel mapping window opens by default.
    ######
    # Check if the channel setup is correct.
    chan_status = pag.locateCenterOnScreen('channels.png', grayscale=grayscale_bool)
    if chan_status == None:
          # Map the channels
          # Channel 1 = ECG
          pag.hotkey('tab')
          pag.hotkey('space')
          pag.typewrite('ECG')
          pag.hotkey('enter')
          # Channel 2 = resp 2 tabs
          for i in range(2): pag.hotkey('tab')
          pag.hotkey('space')
          pag.typewrite('resp')
          pag.hotkey('enter')
          # Channel 3 = Zo 3 tabs
          for i in range(3): pag.hotkey('tab')
          pag.hotkey('space')
          pag.typewrite('zo')
          pag.hotkey('enter')
          # Channel 4 = dZ 4tabs
          for i in range(4): pag.hotkey('tab')
          pag.hotkey('space')
          pag.typewrite('dz')
          pag.hotkey('enter')
    # Go to "OK" button (shift + tab x 5 times)
    for i in range(5): pag.hotkey('shift', 'tab')
    pag.hotkey('enter')

    #===================================================
    # Back to the main window
    #===================================================

    # Check Mode. If mode is not Events, change to events
    pag.click(clickhere)
    for i in range(2): pag.hotkey('tab')
    pag.hotkey('space')
    pag.typewrite('eve')
    pag.hotkey('enter')

    # Check the Auto-Analyze Button, if on. Turn it off
    aa_on = pag.locateCenterOnScreen('aa_on.png',
     grayscale=grayscale_bool, confidence = conf)
    while aa_on != None:
          pag.click(pag.locateCenterOnScreen('aa_on.png', grayscale=grayscale_bool, confidence = conf))
          aa_off = pag.locateCenterOnScreen('auto_analyze_off.png', grayscale=grayscale_bool, , confidence = conf)
    if aa_on != None:
        pag.alert('Cannot turn off the Auto Analyze Button')
        break

    # Check "Filter Events"
    filter_all_status = pag.locateCenterOnScreen('filter_all.png', grayscale=grayscale_bool, , confidence = conf)
    if filter_all_status == None:
        pag.alert('Set Filter Events -> All Events')
        break

    # Check event mode
    event_mode_pre = pag.locateCenterOnScreen('event_mode_pre.png', grayscale=grayscale_bool, , confidence = conf)
    event_mode_post = pag.locateCenterOnScreen('event_mode_post.png', grayscale=grayscale_bool, , confidence = conf)
    [emx, emy] = pag.locateCenterOnScreen('event_mode.png')

    # If UDP markers are present, we need 60 seconds before the FB end marker.
    # If UDP markers are not present, we need 60 seconds before the paced breathing marker

    # Go to the Event Mode
    pag.click(clickhere)
    for i in range(9): pag.hotkey('tab')
    pag.hotkey('space')
    pag.typewrite('pre')
    pag.hotkey('enter')
    # Verify that event mode is Pre
    event_mode_pre = pag.locateCenterOnScreen('event_mode_pre.png', grayscale=grayscale_bool, , confidence = conf)
    if event_mode_pre == None:

        pag.alert("Cannot change the event mode to pre")
        break

    # Set the pre-time to 60
    pag.click(clickhere)
    for i in range(10): pag.hotkey('tab')
    pag.typewrite('60')

    # Set the Event to Use to "User Defined"
    pag.click(clickhere)
    for i in range(11): pag.hotkey('tab')
    pag.hotkey('space')
    pag.typewrite('u')
    pag.hotkey('enter')

    # Event list---
    event_head = pag.locateCenterOnScreen('events_list_head.png', grayscale=grayscale_bool, , confidence = conf)
    pag.click(x = event_head[0], y = event_head[1] + 15)
    for i in range(20): pag.hotkey('up')

    # Set the event (FB end
    fb_end = pag.locateCenterOnScreen('fb_end_inactive.png', grayscale=grayscale_bool, , confidence = conf)
    if fb_end != None:
        pag.click(fb_end)
    elif fb_end == None:
        F2_pb_begin_inactive = pag.locateCenterOnScreen('F2_pb_begin_inactive.png', grayscale=grayscale_bool, , confidence = conf)
        pag.click(F2_pb_begin_inactive)
        if F2_pb_begin_inactive == None:
            continue

##    fb_end_active = pag.locateCenterOnScreen('fb_end_active.png')
##    if fb_end_active == None:
##        pag.alert('Cannot activate Free Breathing End')


    #===================================================
    # Setup 'Impedance Calibration Settings'
    #===================================================

    # Go to the 'ICS Tab'
    pag.click(clickhere)
    for i in range(5): pag.hotkey('tab')
    pag.hotkey('right')

    # Check if the tab is open
    imp_calibration = pag.locateCenterOnScreen('impedance_calibration.png', grayscale=grayscale_bool, confidence = conf)
    if imp_calibration == None:
        pag.alert('Tab is not open. Terminated the process.')
        break

    # Check if Calculate TPR is off
    calculate_TPR = pag.locateCenterOnScreen('calc_TPR_checked.png', grayscale=grayscale_bool, confidence = conf)
    if calculate_TPR != None:
        pag.click(clickhere)
        for i in range(7): pag.hotkey('tab')
        pag.hotkey('space')

    # check 'Electrode Type'
    electrode_type = pag.locateCenterOnScreen('electrode_spot.png', grayscale=grayscale_bool, confidence = conf)
    if electrode_type == None:
        pag.alert('Electrode Type is not Spot')
        break

    # Enter Electrode Distance
    e_distance = pag.locateCenterOnScreen('e_distance.png', grayscale=grayscale_bool, confidence = conf)
    pag.click(x = e_distance[0], y = e_distance[1] + 15, clicks = 2)
    # Write down the electrode cm
    pag.typewrite(str(electrode_cm))

    # Check Rho
##    rho = pag.locateCenterOnScreen('rho.png')
##    if rho == None:
##        rho_txt = pag.locateCenterOnScreen('rho_txt.png')
##        pag.click(x = rho_txt[0], y = rho_txt[1] + 15, clicks = 2)
    pag.hotkey('tab')
    pag.typewrite('135')

    # Change Q point calc method (qcalc_txt_
##    qcalc = pag.locateCenterOnScreen('qcalc_txt.png')
##    pag.click(x = qcalc[0], y = qcalc[1] + 15)
    pag.hotkey('tab')
    pag.hotkey('space')
    pag.typewrite('min')
    pag.hotkey('enter')


    ## Go back to the distance...entering will exit out of the tab
    pag.click(x = e_distance[0], y = e_distance[1] + 15, clicks = 2)

    # Go to b-point calc --> choose max slope
    for i in range(3): pag.hotkey('tab')
    pag.hotkey('space')
    pag.typewrite('max')
    pag.hotkey('enter')

    # Go back s
    pag.click(x = e_distance[0], y = e_distance[1] + 15, clicks = 2)

    # K
    for i in range(4): pag.hotkey('tab')
    pag.typewrite('35')

    # dZdt source
    pag.hotkey('tab')
    pag.hotkey('space')
    pag.typewrite('m')
    pag.hotkey('enter')

    # Go back
    pag.click(x = e_distance[0], y = e_distance[1] + 15, clicks = 2)

    # Ensemble length
    for i in range(6): pag.hotkey('tab')
    pag.typewrite('650')

    # Ensemble Start-R
    pag.hotkey('tab')
    pag.typewrite('200')

    # SV Calc method
    for i in range(2): pag.hotkey('tab')
    pag.hotkey('space')
    pag.typewrite('k')
    pag.hotkey('enter')

    # Go back
    pag.click(x = e_distance[0], y = e_distance[1] + 15, clicks = 2)

    # LVET windowining
    pag.hotkey('shift', 'tab')
    pag.hotkey('space')
    pag.typewrite('f')
    pag.hotkey('enter')

    # Go back
    pag.click(x = e_distance[0], y = e_distance[1] + 15, clicks = 2)

    # LVET Max Offset
    for i in range(2): pag.hotkey('shift', 'tab')
    pag.typewrite('600')

    # Block Size
    pag.hotkey('shift', 'tab')
    pag.typewrite('55')

    # LVET Min Offset
    pag.hotkey('shift', 'tab')
    pag.typewrite('300')





##
##    # Check Ensemble Start-R (estart)
##    estart = pag.locateCenterOnScreen('estart.png')
##    if estart == None:
##        estart_txt = pag.locateCenterOnScreen('estart_txt.png')
##        pag.click(x = estart_txt[0], y = estart_txt[1] + 15, clicks = 2)
##        pag.typewrite('200')
##        pag.click(clickhere)
##
##    # Check Ensemble Length (elength)
##    elength = pag.locateCenterOnScreen('elength.png')
##    if elength == None:
##        estart_txt = pag.locateCenterOnScreen('elength.png')
##        pag.click(x = elength_txt[0], y = elength_txt[1] + 15, clicks = 2)
##        pag.typewrite('650')

    #===================================================
    # Tab: 'R Peak & Artifact Settings'
    #===================================================
    pag.click(clickhere)
    for i in range(5): pag.hotkey('tab')
    pag.hotkey('right')

    # is R peak settings complete?
    r_peak_settings_complete = pag.locateCenterOnScreen('rpeak_settings_complete.png', grayscale=grayscale_bool, confidence = conf)
    if r_peak_settings_complete == None:
        # focus on Max HR
        max_HR = pag.locateCenterOnScreen('max_HR.png', grayscale=grayscale_bool, confidence = conf)
        pag.click(x = max_HR[0], y = max_HR[1] + 15, clicks = 2)
        pag.typewrite('200')

        # Notch Filter
        pag.click(x = max_HR[0], y = max_HR[1] + 15, clicks = 2)
        notch_off = pag.locateCenterOnScreen('notch_off.png', grayscale=grayscale_bool, confidence = conf)
        if notch_off != None:
            pag.hotkey('tab')
            pag.hotkey('space')

        notch_off = pag.locateCenterOnScreen('notch_off.png', grayscale=grayscale_bool, confidence = conf)
        if notch_off != None:
            pag.alert('Cannot turn on the Notch Filter')
            break

        # baseline muscle filter should be checked
        pag.click(x = max_HR[0], y = max_HR[1] + 15, clicks = 2)
        baseline_unchecked = pag.locateCenterOnScreen('baseline_filter_unchecked.png', grayscale=grayscale_bool, confidence = conf)
        if baseline_unchecked != None:
            for i in range(2): pag.hotkey('tab')
            pag.hotkey('space')
            baseline_unchecked = pag.locateCenterOnScreen('baseline_filter_unchecked.png', grayscale=grayscale_bool, confidence = conf)
            if baseline_unchecked != None:
                pag.alert('Cannot turn on the muscle filter')
                break

        # Go back
        pag.click(x = max_HR[0], y = max_HR[1] + 15, clicks = 2)

        # Manual Override Check
        manual_override_on = pag.locateCenterOnScreen('manual_override_on.png', grayscale=grayscale_bool, confidence = conf)
        if manual_override_on != None:
            for i in range(3): pag.hotkey('tab')
            pag.hotkey('space')
            manual_override_on = pag.locateCenterOnScreen('manual_override_on.png', grayscale=grayscale_bool, confidence = conf)
            if manual_override_on != None:
                pag.alert('Cannot turn off the manual override')

        # Go back
        pag.click(x = max_HR[0], y = max_HR[1] + 15, clicks = 2)

        # Notch 60Hz
        notch60 = pag.locateCenterOnScreen('notch60.png', grayscale=grayscale_bool, confidence = conf)
        if notch60 == None:
            for i in range(4): pag.hotkey('tab')
            pag.hotkey('space')
            if notch60 == None:
                pag.alert('Cannot turn on the notch to 60Hz')

        # Go back
        pag.click(x = max_HR[0], y = max_HR[1] + 15, clicks = 2)

        # Peak Detect Sensitivity
        for i in range(5): pag.hotkey('tab')
        pag.typewrite('1')

        # Go back
        pag.click(x = max_HR[0], y = max_HR[1] + 15, clicks = 2)

        # Min HR
        pag.hotkey('shift', 'tab')
        pag.typewrite('40')

        # MAD MED Check On?
        MAD_on = pag.locateCenterOnScreen('MAD_on.png', grayscale=grayscale_bool, confidence = conf)
        if MAD_on == None:
            pag.hotkey('shift', 'tab')
            pag.hotkey('space')
            if MAD_on == None:
                pag.alert('Cannot turn on MAD')
                break

        # Go back
        pag.click(x = max_HR[0], y = max_HR[1] + 15, clicks = 2)

        # IBI Check On?
        IBI_on = pag.locateCenterOnScreen('IBI_on.png', grayscale=grayscale_bool, confidence = conf)
        if IBI_on == None:
            for i in range(3): pag.hotkey('shift', 'tab')
            pag.hotkey('space')
            if IBI_on == None:
                pag.alert('Cannot turn on IBI')
                break

    #===================================================
    # Analyze Window
    #===================================================
    pag.click(clickhere)
    for i in range(7): pag.hotkey('shift', 'tab')
    pag.hotkey('space')

    # Is the EEG inverted
    time.sleep(2)
    pag.hotkey('space')

    # Is the dZ/dt inverted?
    time.sleep(2)
    pag.hotkey('space')

    # Take a screenshot
    shot = pag.screenshot(mwi + '_pre.png')

    # Analyze Window Opens
    write = pag.locateCenterOnScreen('write.png', grayscale=grayscale_bool, confidence = conf)
    pag.click(write)

    # Create a new Excel Output? -> Yes
    pag.hotkey('tab')
    pag.hotkey('space')
    time.sleep(1)

    # Close the analyze window
    pag.hotkey('ctrl', 'w')

    # If debug mode is on, wait for the response.
    if debug == 1:
        proceed = pag.confirm('Do you want to continue?')
        if proceed == 'OK':
            continue
        else:
            pag.alert('The process was stopped by user.')
            break
