CONST HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1 = "DisplayName"
strEntry3 = "VersionMajor"
strEntry4 = "VersionMinor"
SET objReg = GetObject("winmgmts://./root/default:StdRegProv")
objReg.enumKey HKLM, strKey, arrSubkeys
FOR EACH strSubkey IN arrSubkeys
	intRet1 = objReg.getStringValue(HKLM, strKey & strSubkey, strEntry1, strValue1)
	IF intRet1 <> 0 THEN
		objReg.getStringValue HKLM, strKey & strSubkey, strValue1
	END IF
	IF strValue1 <> "" THEN
		name = strValue1
	END IF
	objReg.getDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3
	objReg.getDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4
	IF intValue3 <> "" THEN
		version = intValue3 & "." & intValue4
		complete = 1
	END IF
	
	IF complete = 1 THEN
		ECHO name & vbTab & version
 	END IF
NEXT
