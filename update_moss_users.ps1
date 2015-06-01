[void][reflection.assembly]::Loadwithpartialname("Microsoft.SharePoint") | out-null
[void][reflection.assembly]::Loadwithpartialname("Microsoft.Office.Server.Search") | out-null
[void][reflection.assembly]::Loadwithpartialname("Microsoft.Office.Server") | out-null

# Function:          Get-UserProfileConfigManager
# Description:       return a UserProfileConfigManager object which is used for management of MOSS User Profiles
# Parameters:        PortalURL          URL for the Portal Site Collection    
#
#
function Get-UserProfileConfigManager([string]$PortalURL)
{

# Need to get a PortalContext object 
# as we do not have a HttpContext we need to source one the hard way

$site=new-object Microsoft.SharePoint.SPSite($PortalURL)
$servercontext=[Microsoft.Office.Server.ServerContext]::GetContext($site)
$site.Dispose() # clean up

# Return the UserProfileConfigManager
new-object Microsoft.Office.Server.UserProfiles.UserProfileConfigmanager($servercontext)

}

# Function:           Get-SPProfileManager
# Description:        Return a UserProfileManager object which is used for accessing MOSS User Profiles
# Parameters:         PortalURL          URL for the Portal Site Collection    
#
function Get-SPProfileManager([string]$PortalURL)
{

# Need to get a PortalContext object 
# as we do not have a HttpContext we need to source one the hard way

$site=new-object Microsoft.SharePoint.SPSite($PortalURL)
$servercontext=[Microsoft.Office.Server.ServerContext]::GetContext($site)

$site.Dispose() # clean up

# Return the UserProfileManager
new-object Microsoft.Office.Server.UserProfiles.UserProfileManager($servercontext)


}

# Function:           Update-UserProfileProperty
# Description:        Updates a property for a User in the MOSS User Profiles
function Update-UserProfileProperty()
{
	PARAM
	(
		[string] $siteUrl = $( throw "You must provide a Site Collection Url e.g. 'http://moss/'"),
		[string] $userName = $( throw "You must provide a User Name e.g. 'DOMAIN\USERNAME'"),
		[string] $propName = $( throw "You must provide a User Profile Property Name e.g. 'WorkPhone'"),
		[string] $propValue = $( throw "You must provide a User Profile Property Value e.g. '0400 767 022'")
	)
	END
	{
		if ($propValue -eq "NULL" -or $propValue -eq "" -or $propValue -eq "None")
		{
			Write-Host "Property '$propName' is not set ('$propValue')"
		}
		else
		{
			$cm = get-userprofileconfigmanager $siteUrl 
			$spm = Get-SPProfileManager $siteUrl 
			if ($spm.UserExists($userName))
			{
				$userProfile = $spm.GetUserProfile($userName);
				$tempProp = $spm.Properties.GetPropertyByName($propName);
				if ($tempProp -eq $null)
				{
					throw "User Profile Property '$propName' does not exist!";
				}
				else
				{
					$userProfile[$propName].Value = $propValue;
					$userProfile.Commit();
				}
				Write-Host "'$propName' User Profile Property updated to '$propValue' for '$userName'"
			}
			else
			{		
				Write-Host "User '$userName' does not exist in User Profiles!";	
			}
		}
	}
}

$siteUrl = "https://gateway.manchester.edu";

Import-Csv c:\test.csv | foreach-object {
	$account = "mc\" + $_.sAMAccountName;
	Update-UserProfileProperty $siteUrl $account "PrimaryConstituency" $_.PrimaryConstituency
}