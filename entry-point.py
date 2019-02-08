import pywinauto

from pywinauto.application import Application
app = Application(backend="uia").start('notepad.exe')

# describe the window inside Notepad.exe process
dlg_spec = app.UntitledNotepad
# wait till the window is really open
actionable_dlg = dlg_spec.wait('visible')


#
# If you want to navigate across process boundaries 
# (say Win10 Calculator surprisingly draws its widgets 
# in more than one process) your entry point is a Desktop object.
#

from subprocess import Popen
from pywinauto import Desktop

Popen('calc.exe', shell=True)
dlg = Desktop(backend="uia").Calculator
dlg.wait('visible')

##
#There are many possible criteria 
# for creating window specifications. 
# These are just a few examples.

# can be multi-level
app.window(title_re='.* - Notepad$').window(class_name='Edit')

# can combine criteria
dlg = Desktop(backend="uia").Calculator
dlg.window(auto_id='num8Button', control_type='Button')


# But fortunately pywinauto 
# uses “best match” algorithm to make a lookup 
# resistant to typos and small variations.
app.UntitledNotepad
# is equivalent to
app.window(best_match='UntitledNotepad')

# Unicode characters and special 
# symbols usage is possible through 
# an item access in a dictionary like manner.
app['Untitled - Notepad']
# is the same as
app.window(best_match='Untitled - Notepad')


# By title (window text, name): 
app.Properties.OK.click()
# By title and control type: 
app.Properties.OKButton.click()
# By control type and number: 
app.Properties.Button3.click() # (Note: Button0 and Button1 match the same button, Button2 is the next etc.)
# By top-left label and control type: 
app.OpenDialog.FileNameEdit.set_text("")
# By control type and item text: 
app.Properties.TabControlSharing.select("General")

# More detailed window 
# specification can also be just copied from the method output. Say 
app.Properties.child_window(title="Contains:", auto_id="13087", control_type="Edit")
# To check these names for specified dialog you can use 
app.Properties.print_control_identifiers()

####################################################################
#
#   Second Page
####################################################################

# start()
# is used when the application is not running and you need to start it. 
# Use it in the following way:
app = Application().start(r"c:\path\to\your\application -a -n -y --arguments")


# connect()
# is used when the application to be automated is already launched. 
# To specify an already running application you need to specify one of the following:

# process:
# the process id of the application, e.g.
app = Application().connect(process=2341)
# handle:
# The windows handle of a window of the application, e.g.
app = Application().connect(handle=0x010f0c)
# path:
# The path of the executable of the process 
# (GetModuleFileNameEx is used to find the path of each process and compared 
# against the value passed in) e.g.
app = Application().connect(path=r"c:\windows\system32\notepad.exe")

# There are many different ways of doing this. 
# The most common will be using item or attribute access to select a 
# dialog based on it’s title. e.g
dlg = app.Notepad
dlg = app['Notepad']
# This will return the window that has the highest Z-Order 
# of the top-level windows of the application.
dlg = app.top_window()

# If this is not enough control then you can use the same parameters
# as can be passed to findwindows.find_windows() e.g.
dlg = app.window(title_re="Page Setup", class_name="#32770")
# Finally to have the most control you can use
dialogs = app.windows()

# this will return a list of all the visible, enabled, top level windows of the 
# application. You can then use some of the methods in handleprops 
# module select the dialog you want. 
# Once you have the handle you need then use
#app.window(handle=win)

# If the title of the dialog is very long - then attribute 
# access might be very long to type, in those cases it is usually easier to use
app.window(title_re=".*Part of Title.*")

# There are a number of ways to specify a control, the simplest are
app.dlg.control
app['dlg']['control']

# The 2nd is better for non English OS’s where you need to 
# pass unicode strings e.g. 
app[u'your dlg title'][u'your ctrl title']

#
# Often, when you click/right click on an icon, you get a popup menu. 
# The thing to remember at this point is that the popup menu is a part 
# of the application being automated not part of explorer.

# connect to outlook
#outlook = Application.connect(path='outlook.exe')

# click on Outlook's icon
#taskbar.ClickSystemTrayIcon("Microsoft Outlook")

# Select an item in the popup menu
#outlook.PopupMenu.Menu().get_menu_path("Cancel Server Request")[0].click()

app.wait_cpu_usage_lower(threshold=5) # wait until CPU usage is lower than 5%

app.SendKeys('^a^c') # select all (Ctrl+A) and copy to clipboard (Ctrl+C)
app.SendKeys('+{INS}') # insert from clipboard (Shift+Ins)
app.SendKeys('%{F4}') # close an active window with Alt+F4

# {SCROLLLOCK}, {VK_SPACE}, {VK_LSHIFT}, {VK_PAUSE}, {VK_MODECHANGE},
# {BACK}, {VK_HOME}, {F23}, {F22}, {F21}, {F20}, {VK_HANGEUL}, {VK_KANJI},
# {VK_RIGHT}, {BS}, {HOME}, {VK_F4}, {VK_ACCEPT}, {VK_F18}, {VK_SNAPSHOT},
# {VK_PA1}, {VK_NONAME}, {VK_LCONTROL}, {ZOOM}, {VK_ATTN}, {VK_F10}, {VK_F22},
# {VK_F23}, {VK_F20}, {VK_F21}, {VK_SCROLL}, {TAB}, {VK_F11}, {VK_END},
# {LEFT}, {VK_UP}, {NUMLOCK}, {VK_APPS}, {PGUP}, {VK_F8}, {VK_CONTROL},
# {VK_LEFT}, {PRTSC}, {VK_NUMPAD4}, {CAPSLOCK}, {VK_CONVERT}, {VK_PROCESSKEY},
# {ENTER}, {VK_SEPARATOR}, {VK_RWIN}, {VK_LMENU}, {VK_NEXT}, {F1}, {F2},
# {F3}, {F4}, {F5}, {F6}, {F7}, {F8}, {F9}, {VK_ADD}, {VK_RCONTROL},
# {VK_RETURN}, {BREAK}, {VK_NUMPAD9}, {VK_NUMPAD8}, {RWIN}, {VK_KANA},
# {PGDN}, {VK_NUMPAD3}, {DEL}, {VK_NUMPAD1}, {VK_NUMPAD0}, {VK_NUMPAD7},
# {VK_NUMPAD6}, {VK_NUMPAD5}, {DELETE}, {VK_PRIOR}, {VK_SUBTRACT}, {HELP},
# {VK_PRINT}, {VK_BACK}, {CAP}, {VK_RBUTTON}, {VK_RSHIFT}, {VK_LWIN}, {DOWN},
# {VK_HELP}, {VK_NONCONVERT}, {BACKSPACE}, {VK_SELECT}, {VK_TAB}, {VK_HANJA},
# {VK_NUMPAD2}, {INSERT}, {VK_F9}, {VK_DECIMAL}, {VK_FINAL}, {VK_EXSEL},
# {RMENU}, {VK_F3}, {VK_F2}, {VK_F1}, {VK_F7}, {VK_F6}, {VK_F5}, {VK_CRSEL},
# {VK_SHIFT}, {VK_EREOF}, {VK_CANCEL}, {VK_DELETE}, {VK_HANGUL}, {VK_MBUTTON},
# {VK_NUMLOCK}, {VK_CLEAR}, {END}, {VK_MENU}, {SPACE}, {BKSP}, {VK_INSERT},
# {F18}, {F19}, {ESC}, {VK_MULTIPLY}, {F12}, {F13}, {F10}, {F11}, {F16},
# {F17}, {F14}, {F15}, {F24}, {RIGHT}, {VK_F24}, {VK_CAPITAL}, {VK_LBUTTON},
# {VK_OEM_CLEAR}, {VK_ESCAPE}, {UP}, {VK_DIVIDE}, {INS}, {VK_JUNJA},
# {VK_F19}, {VK_EXECUTE}, {VK_PLAY}, {VK_RMENU}, {VK_F13}, {VK_F12}, {LWIN},
# {VK_DOWN}, {VK_F17}, {VK_F16}, {VK_F15}, {VK_F14}



# try:
#     # wait a maximum of 10.5 seconds for the
#     # window to be found in increments of .5 of a second.
#     # P.int a message and re-raise the original exception if never found.
#     app.wait_until_passes(10.5, .5, self.Exists, (ElementNotFoundError))
# except TimeoutError as e:
#     print("timed out")
#     raise e