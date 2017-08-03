#!python3
import time, pyautogui, pyperclip

orderNo=2318
global orderNo

lastOrder=orderNo+50
global lastOrder

sl=1
global sl

slL=10
global slL
def page():
    while orderNo!=lastOrder:
        pyperclip.copy(str(orderNo))
        pyautogui.click(114, 15)
        time.sleep(sl)
        pyautogui.hotkey('ctrl', 'g')
        time.sleep(sl)
        pyautogui.typewrite(str(orderNo), 0.1)
        time.sleep(sl)
        try:
            center=pyautogui.center(pyautogui.locateOnScreen('submit.png'))
        except:
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.moveTo(659, 271, duration=0.2)
            pyautogui.scroll(400)
            continue
        pyautogui.moveTo(center)
        pyautogui.moveRel(0, 7, 0.2)
        pyautogui.rightClick()
        time.sleep(sl)
        pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('newtab.png')))
        time.sleep(slL)
        pyautogui.click(297, 11)
        time.sleep(sl)
        pyautogui.click(1245, 415)
        time.sleep(sl)
        pyautogui.dragRel(-252, 0, sl, button='left')
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(sl)
        pyautogui.click(380, 15)
        time.sleep(sl)
        pyautogui.click(510, 745)
        time.sleep(sl)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        time.sleep(sl)
        pyautogui.click(468, 748)
        orderNo+=1
for i in range(5):
    page()
    lastOrder+=50
    pyautogui.click(1024, 91)
    time.sleep(slL)
