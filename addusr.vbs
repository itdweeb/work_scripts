Option Explicit

' AD specific Dim
Dim objContainer, objGroup, objRootLDAP, objUser, strCN, strDNSDomain, strFirst, strGroup, strLast, strOU, strPWD, strSam, strUPN, strUser
' File system specific Dim
Dim intRow, objExcel, objFile, objFSO, objShell, objSpread, strDirectory, strFile, strSheet, strWritePath, textFile
' Email specific Dim
Dim objMessage
' Random numbers Dim
Dim intRand, strRand
Const OPEN_FILE_FOR_WRITING = 2
Const ForReading = 1

' Basic setup mess
strOU = "OU=Guest_OU ,"
strSheet = "H:\Scripts\account.xls"
strGroup = "CN=WIFIGUEST"
strFile = "passwords-" & Month(Now) & "-" & Day(Now) & "-" & Hour(Now) & "-" & Minute(Now) & ".txt"
strWritePath = "H:\Scripts\" & strFile
strDirectory = "H:\Scripts\"

' Bind to Active Directory, Guest User container.
Set objRootLDAP = GetObject("LDAP://rootDSE")
Set objContainer = GetObject("LDAP://OU=Guest_OU,DC=manchester,DC=edu")
strDNSDomain = objRootLDAP.Get("defaultNamingContext")
set objGroup = GetObject("LDAP://CN=WIFIGUEST,OU=Groups,OU=MC Users,DC=manchester,DC=edu")

' Open the Excel spreadsheet
Set objExcel = CreateObject("Excel.Application")
Set objSpread = objExcel.Workbooks.Open(strSheet)
intRow = 2 'Row 1 often contains headings

' Create passwords-(DATE).txt
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
objFile = ""
Set textFile = objFSO.OpenTextFile(strWritePath, OPEN_FILE_FOR_WRITING)

' Here is the 'DO...Loop' that cycles through the cells
' Note intRow, x must correspond to the column in strSheet
Do Until objExcel.Cells(intRow,1).Value = ""
   ' Generate random number to append to strCN, strSam, and strUPN to increase uniqueness
   Randomize Timer
   intRand = ((100 * Rnd()) + 10) 
   strRand = Left(CStr(intRand),2)
   
   ' Prepend g_ to further guarantee uniqueness, as these are just temporary accounts
   strFirst = Trim(objExcel.Cells(intRow, 1).Value) 
   strLast = Trim(objExcel.Cells(intRow, 2).Value)
   strCN = "g_" & Left(strFirst,1) & strLast & strRand
   strSam = strCN
   strUPN = strCN & "@manchester.edu"

   ' Generate Unique Password
   strPWD = Left(strFirst,1) & Left(strLast,1) & Month(Now) & Day(Now) & Year(Now)
   
   ' Build the actual User from data in strSheet.
   Set objUser = objContainer.Create("User", "cn=" & strCN)
   objUser.sAMAccountName = strSam
   objUser.userPrincipalName = strUPN
   objUser.givenName = strFirst
   objUser.sn = strLast
   objUser.SetInfo

   ' Separate section to enable account with its password
   objUser.userAccountControl = 512
   objUser.SetPassword strPWD
   objUser.AccountExpirationDate = DateAdd("d", 18, Now)
   objUser.SetInfo

   ' Add user to group
   objGroup.add(objUser.ADsPath)
   
   ' Write username and password to txt file
   textFile.WriteLine(strCN & "				" & strPWD)

intRow = intRow + 1
Loop

objExcel.Quit

' Send email - Remember to change TextBody to reflect location of file, and email addresses accordingly
Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "Chris Schott has created account(s)"
objMessage.From = "ConferenceServices@manchester.edu"
objMessage.To = "serveradmins@manchester.edu"
objMessage.TextBody = "Click here to see what accounts were created: \\triton\staff\cschott\scripts\" & strFile
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.manchester.edu"
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Update
objMessage.Send

' Informs user of completion
MsgBox("Completed adding users!")

' Call delusr.vbs to delete expired accounts
set objShell = WScript.CreateObject("WScript.shell")
objShell.Run "delusr.vbs"

WScript.Quit