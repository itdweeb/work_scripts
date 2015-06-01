Import-Module ActiveDirectory
 $date = [DateTime]::Today.AddDays(-30)
 Get-ADComputer -Filter  ‘PasswordLastSet -le $date’ -properties PasswordLastSet | Sort-Object passwordLastSet | Format-Table -AutoSize dNSHostName,passwordLastSet | Out-File c:\users\cschott\desktop\old-computers.txt
 