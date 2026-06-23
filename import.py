import win32com.client

wps = win32com.client.Dispatch("kwps.Application")
print("WPS COM 版本:", wps.Version)
wps.Quit()