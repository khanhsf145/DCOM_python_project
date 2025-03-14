import win32com.client

dcom_obj = win32com.client.Dispatch("DCOM.Server")
print(dcom_obj.hello())
print(dcom_obj.fetch_user(1))
print(dcom_obj.fetch_request(1))