ComputerName = InputBox("Enter the name of the computer you wish to query")
who = "winmgmts:{impersonationLevel=impersonate}!//"& ComputerName &""
Set Users = GetObject( who ).InstancesOf ("Win32_ComputerSystem")
for each User in Users
MsgBox "The user name for the specified computer is: " & User.UserName
Next