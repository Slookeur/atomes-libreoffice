import uno
import unohelper
from com.sun.star.util import XModifyListener

class MyListener(unohelper.Base, XModifyListener):
    def modified(self, event):
        print("modified", event.Source)
    def disposing(self, event):
        print("disposing")

def test():
    import sys
    try:
        import socket
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.getCurrentComponent()
        listener = MyListener()
        doc.addModifyListener(listener)
        print("Listener added")
    except Exception as e:
        print("Error:", e)

test()
