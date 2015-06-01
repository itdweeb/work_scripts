$90Days = (get-date).adddays(-90)

Get-ADUser -filter {(lastlogondate -notlike "*" -OR lastlogondate -le $90days) -AND (passwordlastset -le $90days) -AND (enabled -eq $True)} -Properties lastlogondate, passwordlastset | Select-Object name, lastlogondate, passwordlastset