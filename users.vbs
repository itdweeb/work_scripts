'This script will list all enabled users on your domain

Const ADS_SCOPE_SUBTREE = 2
Const OPEN_FILE_FOR_WRITING = 2
Const ForReading = 1

Wscript.Echo "The output will be written to your desktop."

strFile = "users.txt"
'clean up this part (dynamic desktop (username env variable))
strWritePath = "C:\Users\cschott\Desktop\" & strFile
strDirectory = "C:\Users\cschott\Desktop\"

Set objFSO1 = CreateObject("Scripting.FileSystemObject")

If objFSO1.FileExists(strWritePath) Then
    Set objFolder = objFSO1.GetFile(strWritePath)

Else
    Set objFile = objFSO1.CreateTextFile(strDirectory & strFile)
    objFile = ""

End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set textFile = fso.OpenTextFile(strWritePath, OPEN_FILE_FOR_WRITING)

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCOmmand.ActiveConnection = objConnection
'sAMAccountType='805306368 defines a normal user account
'805306369 defines a machine account
objCommand.CommandText = "Select Name, userAccountControl from 'LDAP://DC=manchester,DC=edu' Where sAMAccountType='805306368'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    '514 is the decimal value for disabled users
	IF objRecordSet.Fields("userAccountControl").Value <> 514 THEN
		textFile.WriteLine(objRecordSet.Fields("Name").Value)
	END IF
    objRecordSet.MoveNext
Loop

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = Wscript.Arguments
Set objTextFile = objFSO.OpenTextFile(strWritePath, ForReading)

Do Until objTextFile.AtEndOfStream
    strReg = objTextFile.Readline
Loop

WScript.Echo "All done!" 