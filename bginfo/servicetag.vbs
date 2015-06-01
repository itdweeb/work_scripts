SET objWMIService = getObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
SET colSMBIOS = objWMIService.execQuery("Select * from Win32_SystemEnclosure")
FOR EACH objSMBIOS IN colSMBIOS
	ECHO objSMBIOS.serialNumber
NEXT
