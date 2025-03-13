import pythoncom
import win32com.server.util
import win32api
import logging

class MyDCOMServer:
    _public_methods_ = ['Hello', 'AddNumbers','Test']
    _reg_progid_ = "TestDCOM.Server"
    _reg_clsid_ = "{E8CABC68-3C3A-4F4D-8B25-65853ABEE831}"  # Unique CLSID

    def Hello(self):
        return "Hello from Python DCOM Server!"

    def AddNumbers(self, a, b):
        return str(a + b)
    def Test(self):
        return "hello"

if __name__ == "__main__":
    from win32com.server.register import UseCommandLine
    UseCommandLine(MyDCOMServer)
