# coding: utf-8
import win32api
import win32con
import pyautogui
import time

from selenium import webdriver

moveToX = 300
moveToY = 460
menu_num_clicks = 1
secs_between_clicks = 0

pyautogui.click(x=moveToX, y=moveToY, clicks=menu_num_clicks, button='left')
#pyautogui.click(x=moveToX, y=moveToY, clicks=menu_num_clicks, interval=secs_between_clicks, button='left')

print('Press Ctrl-C to quit.')
try:
    while True:
        x, y = pyautogui.position()
        positionStr = 'X: ' + str(x).rjust(4) + ' Y: ' + str(y).rjust(4)
        print(positionStr, end='')
        print('\b' * len(positionStr), end='', flush=True)
except KeyboardInterrupt:
    print('\n')

# The screen resolution size is returned by the size() function as a tuple of two integers. 
# The current X and Y coordinates of the mouse cursor are returned by the position() function.
pyautogui.size()
(1920, 1080)
pyautogui.position()
(187, 567)


### pyautogui.moveTo(100, 200)   # moves mouse to X of 100, Y of 200.
### pyautogui.moveTo(None, 500)  # moves mouse to X of 100, Y of 500.
### pyautogui.moveTo(600, None)  # moves mouse to X of 600, Y of 500.

### pyautogui.moveTo(100, 200)   # moves mouse to X of 100, Y of 200.
### pyautogui.moveRel(0, 50)     # move the mouse down 50 pixels.
### pyautogui.moveRel(-30, 0)     # move the mouse left 30 pixels.
### pyautogui.moveRel(-30, None)  # move the mouse left 30 pixels.

### pyautogui.dragTo(100, 200, button='left')     # drag mouse to X of 100, Y of 200 while holding down left mouse button
### pyautogui.dragTo(300, 400, 2, button='left')  # drag mouse to X of 300, Y of 400 over 2 seconds while holding down left mouse button
### pyautogui.dragRel(30, 0, 2, button='right')   # drag the mouse left 30 pixels over 2 seconds while holding down the right mouse button

### pyautogui.click(x=100, y=200)  # move to 100, 200, then click the left mouse button.
### pyautogui.click(button='right')  # right-click the mouse

### pyautogui.click(x=moveToX, y=moveToY, clicks=menu_num_clicks, button='left')
#pyautogui.click(x=moveToX, y=moveToY, clicks=menu_num_clicks, interval=secs_between_clicks, button='left')



### pyautogui.click(clicks=2)  # double-click the left mouse button
### pyautogui.click(clicks=2, interval=0.25)  # double-click the left mouse button, but with a quarter second pause in between clicks
### pyautogui.click(button='right', clicks=3, interval=0.25)  ## triple-click the right mouse button with a quarter second pause in between clicks

### pyautogui.doubleClick()  # perform a left-button double click

### pyautogui.mouseDown(); pyautogui.mouseUp()  # does the same thing as a left-button mouse click
### pyautogui.mouseDown(button='right')  # press the right button down
### pyautogui.mouseUp(button='right', x=100, y=200)  # move the mouse to 100, 200, then release the right button up.

### pyautogui.scroll(10)   # scroll up 10 "clicks"
### pyautogui.scroll(-10)  # scroll down 10 "clicks"
### pyautogui.scroll(10, x=100, y=100)  # move mouse cursor to 100, 200, then scroll up 10 "clicks"

