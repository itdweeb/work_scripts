SET dtmInstallDate = createObject("wbemScripting.sWbemDateTime")
SET objWMIService = getObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
SET colOperatingSystems = objWMIService.execQuery("Select * from Win32_OperatingSystem")
FOR EACH objOperatingSystem IN colOperatingSystems
	ECHO getmydat(objOperatingSystem.installDate)
NEXT

FUNCTION getmydat(wmitime)
	dtmInstallDate.Value = wmitime
	getmydat = dtmInstallDate.getVarDate
END FUNCTION
