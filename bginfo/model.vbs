winmgt = "winmgmts:{impersonationLevel=impersonate}!//"
SET oWMI_Qeury_Result = getObject(winmgt).instancesOf("Win32_ComputerSystem")
FOR EACH oItem IN oWMI_Qeury_Result
	SET oComputer = oItem
	EXIT FOR
NEXT
IF isNull(oComputer.model) THEN
	sComputerModel = "*no-name* model"
ELSE
	sComputerModel = oComputer.model
END IF
IF isNull(oComputer.manufacturer) THEN
	sComputerManufacturer = "*no-name* manufacturer"
ELSE
	sComputerManufacturer = oComputer.manufacturer
END IF
sComputer = Trim(sComputerModel) & " by " & Trim(sComputerManufacturer)
ECHO sComputer
