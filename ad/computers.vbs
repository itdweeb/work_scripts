'This script will list all computers on your domain
'Created by C.E. Harden August 16 2006

Const ADS_SCOPE_SUBTREE = 2
Const OPEN_FILE_FOR_WRITING = 2
Const ForReading = 1


Wscript.Echo "The output will be written to your desktop."

strFile = "computers.txt"
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
objCommand.CommandText = _
    "Select Name, Location from 'LDAP://OU=Labs,OU=MC Computers,DC=manchester,DC=edu' " _
        & "Where objectClass='computer'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    'Wscript.Echo "Computer Name: " & objRecordSet.Fields("Name").Value
    textFile.WriteLine(objRecordSet.Fields("Name").Value)
    objRecordSet.MoveNext
Loop

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = Wscript.Arguments
Set objTextFile = objFSO.OpenTextFile(strWritePath, ForReading)

Do Until objTextFile.AtEndOfStream
    strReg = objTextFile.Readline
Loop

WScript.Echo "All done!" 