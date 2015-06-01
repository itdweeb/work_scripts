Import-Module ActiveDirectory
 $date = [DateTime]::Today.AddDays(-30)
 Get-ADUSer -Filter  ‘lastLogon -le $date’ -properties lastLogon | Sort-Object lastLogon | Format-Table -AutoSize sAMAccountName,lastLogon | Out-File c:\users\cschott\desktop\old-users.txt
 