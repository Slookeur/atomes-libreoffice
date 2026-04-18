import uno
import unohelper
from com.sun.star.frame import XDispatchProviderInterceptor
from com.sun.star.frame import XDispatch

class MyDispatch(unohelper.Base, XDispatch):
    def dispatch(self, url, args):
        print("Intercepted dispatch:", url.Complete)
        
    def addStatusListener(self, ctrl, url):
        pass
    def removeStatusListener(self, ctrl, url):
        pass

class MyInterceptor(unohelper.Base, XDispatchProviderInterceptor):
    def __init__(self):
        self.master = None
        self.slave = None
        self.dispatch = MyDispatch()

    def queryDispatch(self, url, target, flags):
        print("queryDispatch:", url.Complete)
        if url.Complete == ".uno:Delete" or url.Complete == ".uno:Cut":
            return self.dispatch
        if self.slave:
            return self.slave.queryDispatch(url, target, flags)
        return None

    def queryDispatches(self, requests):
        res = []
        for req in requests:
            res.append(self.queryDispatch(req.FeatureURL, req.FrameName, req.SearchFlags))
        return tuple(res)

    def getMasterDispatchProvider(self):
        return self.master
    def setMasterDispatchProvider(self, master):
        self.master = master
    def getSlaveDispatchProvider(self):
        return self.slave
    def setSlaveDispatchProvider(self, slave):
        self.slave = slave

def test():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    frame = doc.getCurrentController().getFrame()
    interceptor = MyInterceptor()
    try:
        frame.registerDispatchProviderInterceptor(interceptor)
        print("Interceptor added")
    except Exception as e:
        print("Error:", e)

test()
