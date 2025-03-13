import win32com.client

dcom_obj = win32com.client.Dispatch("TestDCOM.Server")
print(dcom_obj.Hello())
print(dcom_obj.Test())
print(dcom_obj.AddNumbers(10, 20))
