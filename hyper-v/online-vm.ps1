# Prompt for the Hyper-V Server to use
$HyperVServer = Read-Host "Specify the Hyper-V Server to use (enter '.' for the local computer)"
# Get all guest KVP data objects on the server in question
$Kvps = gwmi -namespace root\virtualization Msvm_KvpExchangeComponent -computername $HyperVServer
# Create an empty hashtable
$table = @{}
# Go over each of the guest KVP data objects
foreach ($Kvp in [array] $Kvps)
  {
   # Get the OSName value out of the guest KVP data
   $xml = [xml]($Kvp.GuestIntrinsicExchangeItems | ? {$_ -match "OSName"})
   $entry = $xml.Instance.Property | ?{$_.Name -eq "Data"}
   # Filter out unknown operating systems
   if ($entry.Value) {$value = $entry.Value} else {$value = "Unknown"}
   # Count up the values and store it in a hashtable
   if ($table.ContainsKey($value)) 
      {$table[$value] = $table[$value] + 1 }
   else
      {$table[$value] = 1}
   }
# Display the results in a nicely formated and sorted manner
$table.GetEnumerator() | Sort-Object Name | Format-Table -Autosize
