﻿# Initialize some variables used for counting and for output 
$From = Get-Date "20/11/2011" 
$To = $From.AddDays(1) 
 
[Int64] $intSent = $intRec = 0
[Int64] $intSentSize = $intRecSize = 0
[String] $strEmails = $null 
 
Write-Host "DayOfWeek,Date,Sent,Sent Size,Received,Received Size" -ForegroundColor Yellow 
 
Do 
{ 
    # Start building the variable that will hold the information for the day 
    $strEmails = "$($From.DayOfWeek),$($From.ToShortDateString())," 
 
    $intSent = $intRec = 0 
    (Get-TransportServer) | Get-MessageTrackingLog -ResultSize Unlimited -Start $From -End $To | ForEach { 
        # Sent E-mails 
        If ($_.EventId -eq "RECEIVE" -and $_.Source -eq "STOREDRIVER")
		{
			$intSent++
			$intSentSize += $_.TotalBytes
		}
         
        # Received E-mails 
        If ($_.EventId -eq "DELIVER")
		{
			$intRec++
			$intRecSize += $_.TotalBytes
		}
    } 
 
 	$intSentSize = [Math]::Round($intSentSize/1MB, 0)
	$intRecSize = [Math]::Round($intRecSize/1MB, 0)
 
    # Add the numbers to the $strEmails variable and print the result for the day 
    $strEmails += "$intSent,$intSentSize,$intRec,$intRecSize" 
    $strEmails 
 
    # Increment the From and To by one day 
    $From = $From.AddDays(1) 
    $To = $From.AddDays(1) 
} 
While ($To -lt (Get-Date)) 
#While ($To -lt (Get-Date "01/12/2011"))