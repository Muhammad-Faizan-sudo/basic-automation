#I wrote this code, to check obj.close() is working with obj.open(False,90) or not
from win32com import client
import time
def macro_handler():
    obj = client.Dispatch("SeeShell")
    obj.open(90)
    img_path = r"C:\Users\Dell\Documents\Redbuffer_intern\3.png"
    i = obj.setVariable("path",img_path)
    time.sleep(4)
    obj.close()
macro_handler()

#Result: obj.close() is not working with obj.open(False) or working with obj.Open(True)