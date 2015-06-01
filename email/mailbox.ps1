$dir = read-host "Directory to save the CSV file" 
 
get-mailbox * | 
get-mailboxstatistics | 
where {$_.ObjectClass -eq "Mailbox"} | 
Select DisplayName,TotalItemSize,ItemCount,StorageLimitStatus | 
Sort-Object TotalItemSize -Desc | 
export-csv "$dir\mailbox_size.csv"
