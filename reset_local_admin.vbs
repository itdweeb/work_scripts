On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2
Const OPEN_FILE_FOR_WRITING = 2
Const ForReading = 1

Wscript.Echo "The output will be written to your desktop."

strFile = "-admin-computers.txt"
Set WshShell = Wscript.CreateObject("Wscript.Shell")
strDirectory = "C:\Users\" & WshShell.ExpandEnvironmentStrings("%username%") & "\Desktop\"
strWritePath = strDirectory & strFile

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(strWritePath) Then
    Set objFolder = objFSO.GetFile(strWritePath)

Else
    Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
    objFile = ""

End If

Set fSO = CreateObject("Scripting.FileSystemObject")
Set textFile = fSO.OpenTextFile(strWritePath, OPEN_FILE_FOR_WRITING)

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name, Location from 'LDAP://OU=MC Computers,DC=manchester,DC=edu' " _
        & "Where objectClass='computer'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF

	strComputer = objRecordSet.Fields("Name").Value
	if Ping(strComputer) = True then
		Set objUser = GetObject("WinNT://" & objRecordSet.Fields("Name").Value & "/administrator")
		objUser.SetPassword("m!n@sAn0r")
	Else
		textFile.WriteLine(objRecordSet.Fields("Name").Value)
	end if
	
    objRecordSet.MoveNext
Loop

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = Wscript.Arguments
Set objTextFile = objFSO.OpenTextFile(strWritePath, ForReading)

Do Until objTextFile.AtEndOfStream
    strReg = objTextFile.Readline
Loop

WScript.Echo "All done!" 

Function Ping(strComputer)

    dim objPing, objRetStatus

    set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
      ("select * from Win32_PingStatus where address = '" & strComputer & "'")

    for each objRetStatus in objPing
        if IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode<>0 then
    Ping = False
            'WScript.Echo "Status code is " & objRetStatus.StatusCode
        else
            Ping = True
            'Wscript.Echo "Bytes = " & vbTab & objRetStatus.BufferSize
            'Wscript.Echo "Time (ms) = " & vbTab & objRetStatus.ResponseTime
            'Wscript.Echo "TTL (s) = " & vbTab & objRetStatus.ResponseTimeToLive
        end if
    next
End Function 