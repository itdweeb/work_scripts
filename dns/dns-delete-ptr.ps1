param(
# If no DC is specified, find one
$dc =
[system.directoryservices.activedirectory.domaincontroller]::FindOne((new-object System.DirectoryServices.ActiveDirectory.DirectoryContext 'Domain')).Name,
[bool]$forReal=$true
)
 
$reg = [regex]'[A-Z]';
$recs = @(gwmi -comp $dc -namespace 'root/microsoftdns' 'MicrosoftDNS_PTRType' | ?{$reg.IsMatch($_.RecordData)})
if($recs) {$recs | %{
    $container = $_.ContainerName; # 1.in-addr.arpa
    $record = $_.RecordData; # CompName.contoso.com.
    $lowerRecord = $_.RecordData.tolower()
    $owner = $_.OwnerName; # 4.3.2.1.in-addr.arpa
    $shortOwner = $owner.subString(0,$owner.length - $container.length - 1); # 4.3.2
    if($forReal) {dnscmd $dc /RecordDelete "$container." $shortOwner PTR /f | out-null}
    if($forReal){([wmiclass]"\\$dc\root\MicrosoftDNS:MicrosoftDNS_PTRType").CreateInstanceFromPropertyData($DNSServer,$container,$owner,$null,$null,$lowerRecord) | out-null}
    $ret = new-object object;
    $ret | add-member NoteProperty 'InArpa' $owner;
    $ret | add-member NoteProperty 'From' $record;
    $ret | add-member NoteProperty 'To' $lowerRecord;
    $ret | add-member NoteProperty 'ForReal' $forReal;
    $ret
} }