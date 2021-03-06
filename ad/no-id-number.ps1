Import-Module ActiveDirectory
Get-ADUser -filter * -searchbase "ou=Employees,ou=MC Users,dc=manchester,dc=edu" -Properties employeeNumber | Sort-Object employeeNumber | Format-Table samAccountName,employeeNumber -A | Out-File -filepath C:\employee-id.txt
Get-ADUser -filter * -searchbase "ou=Applicants,dc=manchester,dc=edu" -Properties employeeNumber | Sort-Object employeeNumber | Format-Table samAccountName,employeeNumber -A | Out-File -filepath C:\applicant-id.txt
Get-ADUser -filter * -searchbase "ou=Students,ou=MC Users,dc=manchester,dc=edu" -Properties employeeNumber | Sort-Object employeeNumber | Format-Table samAccountName,employeeNumber -A | Out-File -filepath C:\student-id.txt
Get-ADUser -filter * -searchbase "ou=FPAdmits,dc=manchester,dc=edu" -Properties employeeNumber | Sort-Object employeeNumber | Format-Table samAccountName,employeeNumber -A | Out-File -filepath C:\fpadmit-id.txt
