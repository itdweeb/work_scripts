# Prompt for the Hyper-V Server to use
$HyperVServer = Read-Host "Specify the Hyper-V Server to use (enter '.' for the local computer)"
# Get all virtual machine objects on the server in question
$VMs = gwmi -namespace root\virtualization Msvm_ComputerSystem -computername $HyperVServer -filter "Caption = 'Virtual Machine'" 
# Create an empty hashtable
$table = @{}
# Go over each of the virtual machines
foreach ($VM in [array] $VMs) 
  {
   # Get the KVP Object
   $query = "Associators of {$VM} Where AssocClass=Msvm_SystemDevice ResultClass=Msvm_KvpExchangeComponent"
   $Kvp = gwmi -namespace root\virtualization -query $query -computername $HyperVServer
   # Get the OSName value out of the guest KVP data
   $xml = [xml]($Kvp.GuestIntrinsicExchangeItems | ? {$_ -match "OSName"})
   $entry = $xml.Instance.Property | ?{$_.Name -eq "Data"}
   # Filter out offline virtual machines and virtual machines which did not return KVP data
   if ($entry.Value) 
      {$value = $entry.Value}
   elseif ($VM.EnabledState -ne 2)
      {$value = "Offline"}
   else {$value = "Unknown"}
   # Count up the values and store it in a hashtable
   if ($table.ContainsKey($value)) 
      {$table[$value] = $table[$value] + 1 }
   else
      {$table[$value] = 1}
   }
# Display the results in a nicely formated and sorted manner
$table.GetEnumerator() | Sort-Object Name | Format-Table -Autosize