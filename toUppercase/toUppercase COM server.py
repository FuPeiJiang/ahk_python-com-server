class BasicServer:
    # list of all method names exposed to COM
    _public_methods_ = ["toUppercase"]

    @staticmethod
    def toUppercase(string):
        return string.upper()
        
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Error: need to supply arg (""--register"" or ""--unregister"")")
        sys.exit(1)
    else:
        import win32com.server.register
        import win32com.server.exception

        # this server's CLSID
        # NEVER copy the following ID 
        # Use "print(pythoncom.CreateGuid())" to make a new one.
        myClsid="{C70F3BF7-2947-4F87-B31E-9F5B8B13D24F}"
        # this server's (user-friendly) program ID
        myProgID="Python.stringUppercaser"
        
        import ctypes
        def make_sure_is_admin():
            try:
                if ctypes.windll.shell32.IsUserAnAdmin():
                    return
            except:
                pass
            exit("YOU MUST RUN THIS AS ADMIN")
        
        if sys.argv[1] == "--register":
            make_sure_is_admin()
                
            import pythoncom
            import os.path
            realPath = os.path.realpath(__file__)
            dirName = os.path.dirname(realPath)
            nameOfThisFile = os.path.basename(realPath)
            nameNoExt = os.path.splitext(nameOfThisFile)[0]
            # stuff will be written here
            # HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\${myClsid}
            # HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{C70F3BF7-2947-4F87-B31E-9F5B8B13D24F}
            # and here
            # HKEY_LOCAL_MACHINE\SOFTWARE\Classes\${myProgID}
            # HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Python.stringUppercaser
            win32com.server.register.RegisterServer(
                clsid=myClsid,
                # I guess this is {fileNameNoExt}.{className}
                pythonInstString=nameNoExt + ".BasicServer", #toUppercase COM server.BasicServer
                progID=myProgID,
                # optional description
                desc="return uppercased string",
                #we only want the registry key LocalServer32
                #we DO NOT WANT InProcServer32: pythoncom39.dll, NO NO NO
                clsctx=pythoncom.CLSCTX_LOCAL_SERVER,
                #this is needed if this file isn't in PYTHONPATH: it tells regedit which directory this file is located
                #this will write HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{C70F3BF7-2947-4F87-B31E-9F5B8B13D24F}\PythonCOMPath : dirName
                addnPath=dirName,
            )
            print("Registered COM server.")
            # don't use UseCommandLine(), as it will write InProcServer32: pythoncom39.dll
            # win32com.server.register.UseCommandLine(BasicServer)
        elif sys.argv[1] == "--unregister":
            make_sure_is_admin()

            print("Starting to unregister...")

            win32com.server.register.UnregisterServer(myClsid, myProgID)

            print("Unregistered COM server.")
        else:
            print("Error: arg not recognized")
