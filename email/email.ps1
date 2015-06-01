Import-Module ActiveDirectory

Get-ADUser -filter * -properties mail,givenName,sn -SearchBase "OU=MC Users,DC=manchester,DC=edu" | out-file email.csv