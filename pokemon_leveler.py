from pywinauto.application import Application, ProcessNotFoundError
import os


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


gba = connect_to_vba()
dlg = gba.window(title='VisualBoyAdvance')
open_rom(dlg)
set_window(dlg)
