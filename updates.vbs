SET objWMIService = getObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
SET colQuickFixes = objWMIService.execQuery("Select * from Win32_QuickFixEngineering")
FOR EACH objQuickFix IN colQuickFixes
    ECHO objQuickFix.hotFixID
NEXT
