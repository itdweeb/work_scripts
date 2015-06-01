Option Explicit
Dim adoCommand, adoConnection, strBase, strFilter, strAttributes
Dim objRootDSE, strDNSDomain, strQuery, adoRecordset, strMail, strGivenName, strSN, strWritePath, objFile, textFile, objFSO

'Const ADS_SCOPE_SUBTREE = 2
Const OPEN_FILE_FOR_WRITING = 2

strWritePath = "C:\Users\cschott\Desktop\email.csv"
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFile = objFSO.CreateTextFile(strWritePath)
Set textFile = objFSO.OpenTextFile(strWritePath, OPEN_FILE_FOR_WRITING)

' Setup ADO objects.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
Set adoCommand.ActiveConnection = adoConnection

' Search entire Active Directory domain.
Set objRootDSE = GetObject("LDAP://RootDSE")

strDNSDomain = objRootDSE.Get("defaultNamingContext")
strBase = "<LDAP://ou=MC Users," & strDNSDomain & ">"

' Filter on user objects.
strFilter = "(&(objectCategory=person)(objectClass=user))"

' Comma delimited list of attribute values to retrieve.
strAttributes = "mail,givenName,sn"

' Construct the LDAP syntax query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

' Run the query.
Set adoRecordset = adoCommand.Execute

' Enumerate the resulting recordset.
Do Until adoRecordset.EOF
    ' Retrieve values and display.
    strMail = adoRecordset.Fields("mail").Value
    strGivenName = adoRecordset.Fields("givenName").value
	strSN = adoRecordset.Fields("sn").value
    'Wscript.Echo "Email Address: " & strMail & ", First Name: " & strGivenName & ", Last Name: " & strSN
	textFile.WriteLine(strMail & "," & strGivenName & "," & strSN)
    ' Move to the next record in the recordset.
    adoRecordset.MoveNext
Loop

' Clean up.
adoRecordset.Close
adoConnection.Close 