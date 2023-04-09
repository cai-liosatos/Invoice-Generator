# close excel if error occurs between it opening and closing (cant find file error)
# mail setup?
map_update = False
import sys
import ctypes
import views
from views import *
import convertor
from convertor import *

map = views.map
print(views.dates_list)
# quitting python file if main window is forced close
if not map_update: sys.exit()

message = convertor.main(map)
ctypes.windll.user32.MessageBoxW(0, message, 1)