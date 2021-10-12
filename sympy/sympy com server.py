from sympy import simplify, Number, N
from sympy.parsing.sympy_parser import standard_transformations, implicit_multiplication_application, convert_xor
from sympy.parsing.sympy_parser import parse_expr

from decimal import Decimal

from winsound import MessageBeep

transformations = standard_transformations + (implicit_multiplication_application, convert_xor)

def removeTrailingZerosFromNum(num):
    dec = Decimal(str(num))

    tup = dec.as_tuple()
    delta = len(tup.digits) + tup.exponent
    digits = ''.join(str(d) for d in tup.digits)
    if delta <= 0:
        zeros = abs(tup.exponent) - len(tup.digits)
        val = '0.' + ('0' * zeros) + digits
    else:
        val = digits[:delta] + ('0' * tup.exponent) + '.' + digits[delta:]
    val = val.rstrip('0')
    if val[-1] == '.':
        val = val[:-1]
    if tup.sign:
        return '-' + val
    return val


def removeTrailingZerosFromExpr(operatorObject):
    if operatorObject.args:
        return type(operatorObject)(*[removeTrailingZerosFromExpr(i) for i in operatorObject.args])
    else:
        try:
            return Number(removeTrailingZerosFromNum(operatorObject))
        except:
            return operatorObject

def removeTrailingZerosFromExprOrNumber(operatorObject):
    try:
        return removeTrailingZerosFromNum(operatorObject)
    except:
        return removeTrailingZerosFromExpr(operatorObject)

class BasicServer:
    # list of all method names exposed to COM
    _public_methods_ = ["parExprN"]

    @staticmethod
    def parExprN(clipBak):
        parsed = parse_expr(clipBak, transformations=transformations)
        simplified = simplify(N(parsed))
        finalStr = str(removeTrailingZerosFromExprOrNumber(simplified))
        MessageBeep(-1)
        return finalStr.replace("**", "^")

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
        myClsid="{4530C817-6C66-46C8-8FB0-E606970A8DF6}"
        # this server's (user-friendly) program ID, can be anything you want
        myProgID="Python.SimplifyExpr"
        
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
            # HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{c2467d33-71c5-4057-977c-e847c2286882}
            # and here
            # HKEY_LOCAL_MACHINE\SOFTWARE\Classes\${myProgID}
            # HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Python.SimplifyExpr
            win32com.server.register.RegisterServer(
                clsid=myClsid,
                # I guess this is {fileNameNoExt}.{className}
                pythonInstString=nameNoExt + ".BasicServer", #sympy com server.BasicServer
                progID=myProgID,
                # optional description
                desc="(math) SymPy simplify Expression",
                #we only want the registry key LocalServer32
                #we DO NOT WANT InProcServer32: pythoncom39.dll, NO NO NO
                clsctx=pythoncom.CLSCTX_LOCAL_SERVER,
                #this is needed if this file isn't in PYTHONPATH: it tells regedit which directory this file is located
                #this will write HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{4530C817-6C66-46C8-8FB0-E606970A8DF6}\PythonCOMPath : dirName
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
