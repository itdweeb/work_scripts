$NICs = Get-WMIObject Win32_NetworkAdapterConfiguration | where{$_.IPEnabled -eq "TRUE"}
Foreach($NIC in $NICs) 
{
    $IPAddr = read-host "Enter an IP Address"
    $Netmask = read-host "Enter subnet"
    $Gateway = read-host "Enter gateway"
    $DNS1 = read-host "Enter first DNS server"
    $DNS2 = read-host "Enter second DNS server"
    
    $NIC.EnableStatic($IPAddr, $Netmask)
    $NIC.SetGateways($Gateway)
    $NIC.SetDNSServerSearchOrder($DNS1,$DNS2)
    $NIC.SetDynamicDNSRegistration("TRUE")
}