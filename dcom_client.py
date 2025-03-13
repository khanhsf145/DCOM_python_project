import win32com.client

dcom_obj = win32com.client.Dispatch("DCOM.Server")
print(dcom_obj.Hello())
print(dcom_obj.FetchUsers())
