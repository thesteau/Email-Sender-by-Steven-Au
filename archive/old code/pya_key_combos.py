import pyautogui as pya

# Repository

def control_key_combination(key):
    """Repetitive Control Command"""
    pya.keyDown('ctrl')
    pya.press(key)
    pya.keyUp('ctrl')  

def control_shift_combination(key):
    """Repetitive Control Command"""
    pya.keyDown('ctrl')
    pya.keyDown('shift')
    pya.press(key)
    pya.keyUp('shift')
    pya.keyUp('ctrl') 
