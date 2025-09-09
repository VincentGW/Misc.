import pyperclip
import pyautogui

Excel = (2733, -1317)
OnBase = (163, -1101)
Module = (225, -1085)
Document = (314, -761)
Invoice_Field = (-561, -457)
Supplier_Field = (-561, -199)
Vendor = ""
Invoice = ""
Filename = ""
Saved = {}

def Click(coords):
    pyautogui.sleep(.1); pyautogui.click(coords); pyautogui.sleep(.5)

def DoubleClick(coords):
    pyautogui.sleep(.1); pyautogui.doubleClick(coords); pyautogui.sleep(.5)

def Copy():
    pyautogui.sleep(.05)
    pyautogui.keyDown('ctrl'); pyautogui.press('c'); pyautogui.keyUp('ctrl')
    clip = pyperclip.paste().strip()

    return clip

def Clear():
    pyautogui.sleep(.05)
    pyautogui.keyDown('ctrl'); pyautogui.press('a'); pyautogui.keyUp('ctrl')
    pyautogui.sleep(.05)
    pyautogui.keyDown('backspace'); pyautogui.keyUp('backspace')

def Paste(data):
    pyperclip.copy(data); pyautogui.sleep(.1); pyautogui.keyDown('ctrl'); pyautogui.press('v'); pyautogui.keyUp('ctrl'); pyautogui.sleep(.1)

def Y():
    pyautogui.sleep(.05); pyautogui.press('Y')

def Enter():
    pyautogui.sleep(.2); pyautogui.press('enter')
    
def Esc():
    pyautogui.sleep(.3); pyautogui.keyDown('alt'); pyautogui.press('f4'); pyautogui.keyUp('alt')

def TabForward(n):
    for i in range(n):
        pyautogui.sleep(.2); pyautogui.press('tab')

def TabBack(n):
    for i in range(n):
        pyautogui.sleep(.2); pyautogui.hotkey('shift', 'tab')

def Left(n):
    for i in range(n):
        pyautogui.sleep(.05); pyautogui.press('left')

def Right(n):
    for i in range(n):
        pyautogui.sleep(.05); pyautogui.press('right')

def Down(n):
    for i in range(n):
        pyautogui.sleep(.05); pyautogui.press('down')

Click(Excel)

#Start in the Y column
while True:
    clipboard = Copy()
    if clipboard == 'End' or clipboard == "end":
        break
    if clipboard == "Y" or clipboard == "'Y":
        Down(1)
    else:
        Right(6); clipboard = Copy(); Filename = clipboard.strip()
        Left(15); clipboard = Copy(); Vendor = clipboard.strip()
        Right(1); clipboard = Copy(); Invoice = clipboard.strip()
        if (Vendor, Invoice) not in Saved.items():
            print("Vendor:", Vendor, "Invoice #:", Invoice, " not found - adding to dictionary")
            Saved.update({Vendor: Invoice})
            Click(OnBase)
            Click(Invoice_Field); Clear(); Paste(Invoice)
            Click(Supplier_Field); Clear(); Paste(Vendor)
            Enter()
            print('Search submitted')
            pyautogui.sleep(2)
            DoubleClick(Document)
            pyautogui.sleep(.25); TabForward(3); Enter()
            print('Opening Screen 1...')
            pyautogui.sleep(.5); TabBack(3); Enter()
            print('Opening Screen 2...')
            pyautogui.sleep(3);
            clipboard = '0'
            #print("Clipboard:", clipboard)
            while True:
                pyautogui.sleep(.25)
                TabBack(1); clipboard = Copy()
                #print("Clipboard:", clipboard)
                if clipboard == '70':
                    break
            TabForward(2); Enter()
            print('Opening Screen 3...')
            pyautogui.sleep(1)
            while True:
                pyautogui.press('A')
                pyautogui.keyDown('shift'); pyautogui.press('left'); pyautogui.keyUp('shift')
                clipboard = Copy(); pyautogui.press('backspace')
                #print("Clipboard:", clipboard)
                if clipboard == 'A':
                    break
            Paste(Filename); Enter(); pyautogui.sleep(1)
            #Return to previous screen
            Click(Module); Esc()
            Click(Excel)
            Right(8); Y(); Enter()
        else:
            print("Pair found:", Vendor, Invoice, " - skipping")
            Right(8); Y() ; pyautogui.sleep(.1); Enter()

print("Program ended. Total invoices saved:", len(Saved))

#pyautogui.sleep(3)
#cursor = pyautogui.position()