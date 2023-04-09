# close excel if error occurs between it opening and closing (cant find file error)
# mail setup?
map_update = False
import sys
import ctypes
import views
from views import *
import convertor
from convertor import *
import emails
from emails import *

def text_box(message):
    ctypes.windll.user32.MessageBoxW(0, message, 1)

map = views.map
# quitting python file if main window is forced close
if not map_update: sys.exit()

text_box(convertor.main(map))
text_box("Now making email drafts")
text_box(emails.main())
