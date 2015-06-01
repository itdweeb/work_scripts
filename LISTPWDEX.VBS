CONST SEC_IN_DAY = 86400
CONST ADS_UF_DONT_EXPIRE_PASSWD = &h10000
 
strUser = InputBox("Username is: ")
strOU = InputBox("OU is: ")

SET objUserLDAP = GetObject("LDAP://CN=" & strUser & ",OU=" & strOU & ",OU=MC Users,DC=manchester,DC=edu")
intCurrentValue = objUserLDAP.Get("userAccountControl")

IF intCurrentValue AND ADS_UF_DONT_EXPIRE_PASSWD THEN
    Wscript.Echo "The password does not expire."
ELSE
    dtmValue = objUserLDAP.PasswordLastChanged
    Wscript.Echo "The password was last changed on " & DateValue(dtmValue) & " at " & TimeValue(dtmValue) & VbCrLf & "The difference between when the password was last set and today is " & int(now - dtmValue) & " days"
    intTimeInterval = int(now - dtmValue)
    SET objDomainNT = GetObject("WinNT://MC")
    intMaxPwdAge = objDomainNT.Get("MaxPasswordAge")
    IF intMaxPwdAge < 0 THEN
        WScript.Echo "The Maximum Password Age is set to 0 in the domain. Therefore, the password does not expire."
    ELSE
        intMaxPwdAge = (intMaxPwdAge/SEC_IN_DAY)
        Wscript.Echo "The maximum password age is " & intMaxPwdAge & " days"
        IF intTimeInterval >= intMaxPwdAge THEN
			Wscript.Echo "The password has expired."
        ELSE
			Wscript.Echo "The password will expire on " & DateValue(dtmValue + intMaxPwdAge) & " (" & int((dtmValue + intMaxPwdAge) - now) & " days from today)."
        END IF
    END IF
END IF
