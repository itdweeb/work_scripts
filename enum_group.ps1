$root=([ADSI]"").distinguishedName
$Group = [ADSI]("LDAP://CN=New Students e-forms,OU=Groups,OU=MC Users,"+ $root)
Write-Output $Group.member >> H:\Scripts\members.txt