' This script will delete all expired users in OU=Guest_OU,dc=manchester,dc=edu
Option Explicit

Dim objOU, objUser, objRootDSE, strContainer, strCN, strObject, strDNSDomain

' Bind to Active Directory Domain
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("DefaultNamingContext")
strObject = "user"
strContainer = "OU=Guest_OU, "
strContainer = strContainer & strDNSDomain
set objOU = GetObject("LDAP://" & strContainer)

' 1/1/1970 is the magical date for accounts that don't expire (read:  most accounts)
For each objUser in objOU
   If objUser.accountExpirationDate < Now() AND objUser.accountExpirationDate <> "1/1/1970" then
       strCN = "CN=" & objUser.cn
       objOU.delete strObject, strCN
   End if
next

wscript.echo "Completed deleting expired users!"

wscript.quit