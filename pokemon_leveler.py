import pyautogui
from pywinauto.application import Application, ProcessNotFoundError
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys
from pywinauto.win32structures import RECT
import os
import time
import cv2
import numpy as np
import ctypes
import win32gui

method = cv2.TM_SQDIFF_NORMED

SCREEN_RES_X = 240
SCREEN_RES_Y = 160

CHARACTER_STARTING_POS = (((SCREEN_RES_X // 2) - 8), ((SCREEN_RES_Y // 2) - 8))
CHARACTER_ENDING_POS = (((SCREEN_RES_X // 2) + 8), (SCREEN_RES_Y // 2) + 8)


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
    # 8px padding on each side and 43px for menu on top
    main_dlg.move_window(x=None, y=None, width=SCREEN_RES_X + 16, height=SCREEN_RES_Y + 59, repaint=True)
    main_dlg.set_focus()


# Tests if a template image is in another
# @Param img1: Larger image we are looking for the template in
# @Param img2: Template image we are searching for in the larger image
# @Param threshold: Threshold we would accept a match with
def img_contains(img1, img2, threshold):
    # Execute matchTemplate to see if the img2 is in img1
    match = cv2.matchTemplate(img1, img2, method)
    mn, _, mnLoc, _ = cv2.minMaxLoc(match)
    # If the min value was les than threshold, the template is not in the parent image
    return mn <= threshold


# Presses and releases a key after a delay
# @Param key: Key to be pressed
def press_key(key):
    send_keys("{" + key + " down}")
    time.sleep(.0625)
    send_keys("{" + key + " up}")


def screenshot(window_title=None):
    if window_title:
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
            x, y, x1, y1 = win32gui.GetClientRect(hwnd)
            x, y = win32gui.ClientToScreen(hwnd, (x, y))
            x1, y1 = win32gui.ClientToScreen(hwnd, (x1 - x, y1 - y))
            im = pyautogui.screenshot(region=(x, y, x1, y1))
            return im
        else:
            print('Window not found!')
    else:
        im = pyautogui.screenshot()
        return im


# Checks if there is a save file, if not then it begins a new game
def start_game(main_dlg):
    no_save_result = False
    save_result = False

    send_keys("{SPACE down}")
    # Keep pressing z until we get to the main menu
    while not no_save_result and not save_result:
        press_key('z')
        # Wait for next frame to load before processing image
        time.sleep(1)
        # Capture an image of the screen and convert it to cv2
        img = screenshot(main_dlg.window_text())
        numpy_img = np.array(img)
        parent_img = cv2.cvtColor(numpy_img, cv2.COLOR_RGB2BGR)
        # Read a template image of the main menu
        no_save_img = cv2.imread('img/menu_no_save.PNG')
        save_img = cv2.imread('img/menu_save.PNG')
        # Check if we are at the main menu
        no_save_result = img_contains(parent_img, no_save_img, .003)
        save_result = img_contains(parent_img, save_img, .003)

    send_keys("{SPACE up}")
    # If there is a saved game, press z to get into the game
    if save_result:
        print('Saved game found...')
        press_key('z')
    # If there wasn't a saved game, inform the user and have them retry
    elif no_save_result:
        print('No saved game found. Please ensure you have a .sav file in the same directory as your ROM')
        exit(1)


# Gets 16x16 tiles in each cardinal direction of the player's character from img and crops them into separate images
# @Param img: Parent image to get tiles for
# @return: dictionary of images representing the 4 tiles to the top, left, bottom, and right
def get_directions(img):
    # Crop images in each direction
    top_img = img[CHARACTER_STARTING_POS[1] - 16: CHARACTER_STARTING_POS[1],
              CHARACTER_STARTING_POS[0]: CHARACTER_STARTING_POS[0] + 16].copy()
    left_img = img[CHARACTER_STARTING_POS[1]: CHARACTER_STARTING_POS[1] + 16,
               CHARACTER_STARTING_POS[0] - 16: CHARACTER_STARTING_POS[0]].copy()
    right_img = img[CHARACTER_ENDING_POS[1] - 16: CHARACTER_ENDING_POS[1],
                CHARACTER_ENDING_POS[0]: CHARACTER_ENDING_POS[0] + 16].copy()
    bottom_img = img[CHARACTER_ENDING_POS[1]: CHARACTER_ENDING_POS[1] + 16,
                 CHARACTER_ENDING_POS[0] - 16: CHARACTER_ENDING_POS[0]].copy()
    # cv2.rectangle(img, CHARACTER_STARTING_POS, (CHARACTER_STARTING_POS[0] + 16, CHARACTER_STARTING_POS[1] - 16), (0, 0, 255), 1)
    # cv2.rectangle(img, CHARACTER_STARTING_POS, (CHARACTER_STARTING_POS[0] - 16, CHARACTER_STARTING_POS[1] + 16), (0, 0, 255), 1)
    # cv2.rectangle(img, CHARACTER_ENDING_POS, (CHARACTER_ENDING_POS[0] + 16, CHARACTER_ENDING_POS[1] - 16), (0, 0, 255), 1)
    # cv2.rectangle(img, CHARACTER_ENDING_POS, (CHARACTER_ENDING_POS[0] - 16, CHARACTER_ENDING_POS[1] + 16), (0, 0, 255), 1)
    # cv2.imshow('Top', top_img)
    # cv2.imshow('Left', left_img)
    # cv2.imshow('Bottom', bottom_img)
    # cv2.imshow('Right', right_img)
    # Add images to dictionary and return it
    imgs = {}
    imgs['top'] = top_img
    imgs['left'] = left_img
    imgs['bottom'] = bottom_img
    imgs['right'] = right_img
    return imgs


# Take screenshot of current game. Analyze tiles in all 4 directions to determine if grass is in any of them
def map_image(main_dlg):
    # Screenshot current frame and convert it into an opencv image
    img = screenshot(main_dlg.window_text())
    numpy_img = np.array(img)
    parent_img = cv2.cvtColor(numpy_img, cv2.COLOR_RGB2BGR)
    # Open a template image of grass for matching
    child_img = cv2.imread('img/grass.PNG')
    #cv2.rectangle(parent_img, CHARACTER_STARTING_POS, CHARACTER_ENDING_POS, (0, 0, 255), 1)
    # Call get_directions to get the adjacent tiles
    directions = get_directions(parent_img)
    # cv2.imshow("HERE", parent_img)
    # cv2.waitKey(0)
    results = []
    for direction in directions:
        result = img_contains(directions[direction], child_img, .19)
        print(f'{result} for {direction}')
        results.append(result)


gba = connect_to_vba()
dlg = gba.window(title_re='VisualBoyAdvance.*')
# open_rom(dlg)
set_window(dlg)
# start_game(dlg)
# time.sleep(5)
map_image(dlg)
