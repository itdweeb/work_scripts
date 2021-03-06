$root=([ADSI]"").distinguishedName
$Group = [ADSI]("LDAP://CN=WSUS_Desktop,OU=Computer,OU=Groups,"+ $root)
Write-Output $Group.member >> c:\members.txt
$Group = [ADSI]("LDAP://CN=WSUS_Servers,OU=Computer,OU=Groups,"+ $root)
Write-Output $Group.member >> c:\members.txt
$Group = [ADSI]("LDAP://CN=WSUS_Detect_Only,OU=Computer,OU=Groups,"+ $root)
Write-Output $Group.member >> c:\members.txt
$Group = [ADSI]("LDAP://CN=WSUS_DomainControllers,OU=Computer,OU=Groups,"+ $root)
Write-Output $Group.member >> c:\members.txt
$Group = [ADSI]("LDAP://CN=WSUS_Exchange_Servers,OU=Computer,OU=Groups,"+ $root)
Write-Output $Group.member >> c:\members.txt
$Group = [ADSI]("LDAP://CN=WSUS_MB_ATM,OU=Computer,OU=Groups,"+ $root)
Write-Output $Group.member >> c:\members.txt
$Group = [ADSI]("LDAP://CN=WSUS_Server_Baseline_No_WSUS,OU=Computer,OU=Groups,"+ $root)
Write-Output $Group.member >> c:\members.txt
$Group = [ADSI]("LDAP://CN=WSUS_StandaloneServers,OU=Computer,OU=Groups,"+ $root)
Write-Output $Group.member >> c:\members.txt
Get-Content c:\members.txt | sort > c:\wsus.txt
get-adcomputer -filter * -Properties * | ft distinguishedName | out-file c:\all.txt
get-content c:\all.txt | select-object -skip 2 | sort > c:\computers.txt
