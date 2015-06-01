SET objWMIService = getObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
SET colShares = objWMIService.execQuery("Select * from Win32_Share")
FOR EACH objShare IN colShares
    ECHO objShare.name & vbTab & objShare.path
NEXT
