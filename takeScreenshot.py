from PIL import ImageGrab
import Image, ImageMath
import os
import time
import locationDatas

def screenShot(location):
    if (location == 'Case'):
        bounds = (locationDatas.x_valUpperCase,locationDatas.y_valUpperCase,
                  locationDatas.x_valLowerCase,locationDatas.y_valLowerCase)
    img = ImageGrab.grab(bounds)
    temp = set(img.getdata())
    return temp
    
        

def compareScreenShot(before,after):
    if (before == after):
        return False
    else:
        return True
