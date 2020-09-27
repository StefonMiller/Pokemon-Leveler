from pywinauto.application import Application, ProcessNotFoundError
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys
import os
import time
import cv2
import numpy as np

method = cv2.TM_SQDIFF_NORMED


# Sets up pywinauto Application object. If a window is already open, it connects to that. Otherwise it creates a new
# Application
# @Return: Application object referencing the VBA window
def connect_to_vba():
    try:
        app = Application(backend='win32').connect(path=r'VBA/VisualBoyAdvance.exe')
    except ProcessNotFoundError:
        app = Application(backend='win32').start('VBA/VisualBoyAdvance.exe')
    return app


# Opens the ROM in the ROM folder in the project folder
# @Param main_dlg: Current dialog for the main window
def open_rom(main_dlg):
    # File -> Open to open the menu for selecting a ROM
    main_dlg.menu_select('File->Open...')
    # Get the absolute path of the ROM in the ROM folder
    rom_path = os.path.abspath('ROM/Pokemon - Emerald Version (U).gba')
    # Set the current dialog to the 'Select ROM' window
    rom_dlg = gba.window(title='Select ROM')
    # Get the field where we input the file name
    file_name_field = rom_dlg['Edit:']
    # Set the field text to the absolute path and click the open button
    file_name_field.set_edit_text(rom_path)
    open_button = rom_dlg['&Open']
    open_button.click()


# Sets the VBA window as active and resizes it
# @Param main_dlg: Dialog for the main window
def set_window(main_dlg):
    main_dlg.move_window(x=None, y=None, width=480, height=320, repaint=True)
    main_dlg.set_focus()


# Tests if a template image is in another
# @Param img1: Larger image we are looking for the template in
# @Param img2: Template image we are searching for in the larger image
# @Param threshold: Threshold we would accept a match with
def img_contains(img1, img2, threshold):
    # Execute matchTemplate to see if the 'Continue' option was found
    match = cv2.matchTemplate(img1, img2, method)
    mn, _, mnLoc, _ = cv2.minMaxLoc(match)

    # If the min value was les than .0001, the user does have a saved game
    return mn <= threshold


# Checks if there is a save file, if not then it begins a new game
def start_game(main_dlg):
    matching_debug = False
    print('Waiting for ROM to load...')
    time.sleep(2)

    # Capture an image of the screen and convert it to cv2
    img = main_dlg.capture_as_image()
    numpy_img = np.array(img)
    parent_img = cv2.cvtColor(numpy_img, cv2.COLOR_RGB2BGR)

    # Read a template image of the main menu
    no_save_img = cv2.imread('img/menu_no_save.PNG')
    save_img = cv2.imread('img/menu_save.PNG')
    # Check if we are at the main menu
    no_save_result = img_contains(parent_img, no_save_img, .003)
    save_result = img_contains(parent_img, save_img, .003)

    while not no_save_result and not save_result:
        # Capture an image of the screen and convert it to cv2
        img = main_dlg.capture_as_image()
        numpy_img = np.array(img)
        parent_img = cv2.cvtColor(numpy_img, cv2.COLOR_RGB2BGR)

        # Read a template image of the main menu
        no_save_img = cv2.imread('img/menu_no_save.PNG')
        save_img = cv2.imread('img/menu_save.PNG')
        # Check if we are at the main menu
        no_save_result = img_contains(parent_img, no_save_img, .003)
        save_result = img_contains(parent_img, save_img, .003)
        print('Attempting to type z')
        main_dlg.type_keys('z', vk_packet=False)
        # send_keys("z", vk_packet=False)

    if save_result:
        print('Save menu found')
    elif no_save_result:
        print('You need to start a game and get some pokemon before you can level them!')
        exit(1)


gba = connect_to_vba()
dlg = gba.window(title_re='VisualBoyAdvance.*')
open_rom(dlg)
set_window(dlg)
start_game(dlg)
