import win32com.client as win32


def conexion ():
    
    SapGuiAuto  = win32.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32.CDispatch:
        return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32.CDispatch:
        SapGuiAuto = None
        return
    
    connection = application.Children(0)
    if not type(connection) == win32.CDispatch:
        application = None
        SapGuiAuto = None
        return
        
    session    = connection.Children(0)
    if not type(session) == win32.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return
    return session

def ingresar_transaccion():
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NMEK32"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtRV13A-KSCHL").text = "j3ap"
    session.findById("wnd[0]/usr/ctxtRV13A-KSCHL").caretPosition = 4
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[8,0]").select
    session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[8,0]").setFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


if __name__ == "__main__":
    global session
    session = conexion()
    ingresar_transaccion()


