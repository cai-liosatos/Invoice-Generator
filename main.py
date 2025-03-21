map_update = False
import sys
import ctypes
import views
from views import *
import convertor
from convertor import *
import emails
from emails import *
from ctypes import c_int, WINFUNCTYPE, windll
from ctypes.wintypes import HWND, LPCWSTR, UINT

# ctypes parameters
prototype = WINFUNCTYPE(c_int, HWND, LPCWSTR, LPCWSTR, UINT)
paramflags = (1, "hwnd", 0), (1, "text", "Hi"), (1, "caption", "Lorem ipsum dolor sit amet"), (1, "flags", 0)
MessageBox = prototype(("MessageBoxW", windll.user32), paramflags)

def text_box(message):
    MessageBox(text=message)

# def text_box(message):
#     ctypes.windll.user32.MessageBoxW(0, message, 1)

map = views.map
# quitting python file if main window is forced close
if not map_update: sys.exit()

message = convertor.main(map)
if not message:
    text_box("Successfully created invoices")
    text_box(emails.main())
else:
    text_box(message)
