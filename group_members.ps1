Import-Module ActiveDirectory
Get-ADGroup -Filter {sAMAccountName -like "Web_*"} -Properties member | Out-File groups.txt