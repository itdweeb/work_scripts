Import-Module ActiveDirectory
$siteName =  "MANCHESTER"
$configNCDN = (Get-ADRootDSE).ConfigurationNamingContext
$siteContainerDN = ("CN=Sites," + $configNCDN)
$siteDN = "CN=" + $siteName + "," + $siteContainerDN
$siteObj = Get-ADObject -Identity $siteDN -properties "siteObjectBL", "description", "location" 
foreach ($subnetDN in $siteObj.siteObjectBL) {
    Get-ADObject -Identity $subnetDN -properties "siteObject", "description", "location" | fl -property name
}
