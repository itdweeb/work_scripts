## Requires Exchange Management Shell

###############################################################################
#		Change these variables
$toAddress = "wendy@kmlltdlaw.com" , "mike@kmlltdlaw.com"
$fromAddress = "administrator@kmlltdlaw.com"
$emailServer = "mail.kmlltdlaw.com"
$emailSubject = "Weekly Status Report"
$logPath = "C:\script_logs\"
$archivePath = "C:\script_archives\"
$exchangeLog = "mailbox_report.txt"
$driveLog = "drive_report.txt"
#
###############################################################################

# Gather and generate some basic information
$hostname = (Get-WmiObject Win32_ComputerSystem).Name
$dns = (Get-WmiObject Win32_ComputerSystem).Domain
$fqdn = ($hostname + "." + $dns)
$date = Get-Date
$transport = Get-TransportServer

$emailBody = @"
Attached are the weekly status reports for $($fqdn) for 
</br>
the week of $($date.Month)-$($date.Day).
</br></br>
The reports include the size and item count of all mailboxes in the domain, as 
</br>
well as disk usage statistics for all local hard drives.
</br>
"@

# Find all exchange servers with the mailbox role and get the stats on all Mailbox class objects, so none of the system mailboxes
Get-MailboxServer | Get-MailboxStatistics | Where {$_.ObjectClass -eq "Mailbox"} | Sort-Object TotalItemSize -Descending | Format-Table -AutoSize @{Label = 'Name' ; Expression = {$_.DisplayName}} , @{Label = 'Total Items' ; Expression = {$_.ItemCount}} , @{Label = 'Total Size' ; Expression = {$_.TotalItemSize}} , @{Label = 'Last Accessed' ; Expression = {$_.LastLogonTime}} | Out-File ($logPath + $exchangeLog)

# Get all local drives and calculate used size based on total and free space
Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" | Format-Table -AutoSize @{Label = 'Drive Letter' ; Expression = {$_.DeviceID}} , @{Label = 'Description' ; Expression = {$_.VolumeName}} , @{Label = 'Total Size' ; Expression = {"{0:N2}" -f ($_.Size/1GB)}} , @{Label = 'Free Space' ; Expression = {"{0:N2}" -f ($_.FreeSpace/1GB)}} , @{Label = 'Used Space' ; Expression = {"{0:N2}" -f ($_.Size/1GB) - "{0:N2}" -f ($_.FreeSpace/1GB)}} | Out-File ($logPath + $driveLog)

# Send an email with the generated logs attached
Send-MailMessage -To $toAddress -From $fromAddress -Subject $emailSubject -Body $emailBody -SmtpServer $emailServer -BodyAsHtml -Attachments ($logPath + $exchangeLog) , ($logPath + $driveLog) 

# Get number of emails sent and received per user
Get-MessageTrackingLog -ResultSize unlimited -Start $date.addDays(-7) -End $date -eventid RECEIVE | Where-Object {$_.sender -match "kml" -or "borkanscahill"} | Group-Object -Property Sender | sort-object -property count -Descending | ft -autosize count,name

# Pause to make sure the email sends before renaming and archiving log files
Start-Sleep -Seconds 60

# Rename reports and move to archive folder
Move-Item ($logPath + $exchangeLog) ($archivePath + "$($date.Year)-$($date.Month)-$($date.Day)-$($exchangeLog)")
Move-Item ($logPath + $driveLog) ($archivePath + "$($date.Year)-$($date.Month)-$($date.Day)-$($driveLog)")




possibly move file to some kind of compressed archive
check for folder existence and create if necessary
