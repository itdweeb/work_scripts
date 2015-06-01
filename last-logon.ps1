Import-Module ActiveDirectory
$OU = Read-Host "What OU would you like to scan (full LDAP sytax)?"
Get-ADUser -filter * -SearchBase $OU -Properties 'lastLogon'