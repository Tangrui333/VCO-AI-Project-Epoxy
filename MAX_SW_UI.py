import pygetwindow as gw
import pyautogui
from PIL import ImageGrab
import cv2
import numpy as np
import time
import win32gui
import win32con
import pyscreenshot as ImageGrab
from PIL import Image
import tkinter as tk
from tkinter import messagebox
import os

#find UI window name
def find_window_by_title_part(title_part):
   def callback(hwnd, extra):
       if title_part in win32gui.GetWindowText(hwnd):
           extra.append(hwnd)
   hwnds = []
   win32gui.EnumWindows(callback, hwnds)
   return hwnds

#make sure window status, max/min/normal
def get_window_state(hwnd):
   window_placement = win32gui.GetWindowPlacement(hwnd)
   print(f"Window placement: {window_placement}")  # print detail info of the window
   return window_placement[1]

#max window and show it beyond other windows.
def maximize_and_show_window(hwnd):
   win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
   win32gui.SetForegroundWindow(hwnd)

# cropping the area we are focusing on
def crop_rectangle(image_path1, x, y, z, w):
   img = Image.open(image_path1)
   cropped_img = img.crop((x, y, z, w))
   cropped_img.show()
   # GV value calculation
   grayscale_image = cropped_img.convert('L')  # convert to gray image
   grayscale_array = grayscale_image.getdata()  # to get GV array
   average_grayscale = sum(grayscale_array) / len(grayscale_array)  # 计算平均灰度值
   return average_grayscale

#pops out warning window
def show_alert():
   root = tk.Tk()
   root.withdraw()  # hide main window
   root.attributes('-topmost', True)
   messagebox.showinfo("Warning", "Tray mismatch please esclate to L2/ES")
   root.destroy()


#main script 
window_title_part = 'Calculator'
hwnds = find_window_by_title_part(window_title_part)
if hwnds:
   for hwnd in hwnds:
       window_text = win32gui.GetWindowText(hwnd)
       print(f"got window : '{window_text}'")
       window_state = get_window_state(hwnd)
       if window_state in [win32con.SW_SHOWMINIMIZED, win32con.SW_SHOWNORMAL, win32con.SW_SHOWMAXIMIZED]:
           print(f"Window: '{window_text}' maximized")
           maximize_and_show_window(hwnd)
       else:
           print(f"window '{window_text}' showed beyond all other windows")
   path = r"C:\Users\chaoguo1\XCG1\edge5.jpg"
   stop_x, stop_y = 400 , 600
   im = ImageGrab.grab()
   # im.show()
   im.save(path)  # save image
   if os.path.exists(path):
      x, y, z, w = 1450, 120, 1700, 260  # define the location you want
      average_grayscale = crop_rectangle(path, x, y, z, w)
      print(f"GV mean: {average_grayscale}")
      if average_grayscale < 240:
         pyautogui.moveTo(stop_x,stop_y)
         pyautogui.click()
         time.sleep(1)
         show_alert()
      else:
         pass  
   else:
    print(f"Path:'{path} 'not existed")
else:
   print(f"Didn't find window named '{window_title_part}'")


