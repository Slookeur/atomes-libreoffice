import uno
def test():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    try:
        from com.sun.star.awt.MessageBoxType import QUERYBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_YES_NO
        toolkit = smgr.createInstance("com.sun.star.awt.Toolkit")
        frame = doc.getCurrentController().getFrame()
        peer = frame.getContainerWindow()
        box = toolkit.createMessageBox(peer, QUERYBOX, BUTTONS_YES_NO, "Test", "This is a test. Yes or no?")
        res = box.execute()
        print("Result:", res)
    except Exception as e:
        print("Error:", e)
test()
