﻿<#
.SYNOPSIS
Get-MailboxReport.ps1 - Mailbox report generation script.

.DESCRIPTION 
Generates a report of useful information for
the specified server, database, mailbox or list of mailboxes.
Use only one parameter at a time depending on the scope of
your mailbox report.

.OUTPUTS
Single mailbox reports are output to the console, while all other
reports are output to a CSV file.

.PARAMETER all
Generates a report for all mailboxes in the organization.

.PARAMETER server
Generates a report for all mailboxes on the specified server.

.PARAMETER database
Generates a report for all mailboxes on the specified database.

.PARAMETER file
Generates a report for mailbox names listed in the specified text file.

.PARAMETER mailbox
Generates a report only for the specified mailbox.

.PARAMETER filename
(Optional) Specifies the CSV file name to be used for the report.
If no file name specificed then a unique file name is generated by the script.

.EXAMPLE
.\Get-MailboxReport.ps1 -database HO-MB-01
Returns a report with the mailbox statistics for all mailbox users in
database HO-MB-01

.EXAMPLE
.\Get-MailboxReport.ps1 -file .\users.txt
Returns a report with the mailbox statistics for all mailbox users in
the file users.txt. Text file should contain names in a format that
will work for Get-Mailbox, such as the display name, alias, or primary
SMTP address.

.EXAMPLE
.\Get-MailboxReport.ps1 -server ex2010-mb1
Generates a report with the mailbox statisitcs for all mailbox users
on ex2010-mb1

.EXAMPLE
.\Get-MailboxReport.ps1 -server ex2010-mb1 -filename ex2010-mb1.csv
Generates a report with the mailbox statisitcs for all mailbox users
on ex2010-mb1, and uses the custom file name of ex2010-mb1.csv

.LINK
http://exchangeserverpro.com/powershell-script-create-mailbox-size-report-exchange-server-2010

.NOTES
Written By: Paul Cunningham
Website:	http://exchangeserverpro.com
Twitter:	http://twitter.com/exchservpro

Change Log
V1.0, 2/2/2012 - Initial version
V1.1, 27/2/2012 - Improved recipient scope settings, exception handling, and custom file name parameter.
#>


param(
	[Parameter(ParameterSetName='database')] [string]$database,
	[Parameter(ParameterSetName='file')] [string]$file,
	[Parameter(ParameterSetName='server')] [string]$server,
	[Parameter(ParameterSetName='mailbox')] [string]$mailbox,
	[Parameter(ParameterSetName='all')] [switch]$all,
	[string]$filename
)

#...................................
# Variables
#...................................

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$report = @()


#Set recipient scope
$2007snapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
if ($2007snapin)
{
	$AdminSessionADSettings.ViewEntireForest = 1
}
else
{
	$2010snapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010
	if ($2010snapin)
	{
		Set-ADServerSettings -ViewEntireForest $true
	}
}


#If no filename specified, generate report file name with random strings for uniqueness
#Thanks to @proxb and @chrisbrownie for the help with random string generation

if ($filename)
{
	$reportfile = $filename
}
else
{
	$timestamp = Get-Date -UFormat %Y%m%d-%H%M
	$random = -join(48..57+65..90+97..122 | ForEach-Object {[char]$_} | Get-Random -Count 6)
	$reportfile = "MailboxReport-$timestamp-$random.csv"
}


#...................................
# Script
#...................................

#Add dependencies
Import-Module ActiveDirectory

#Get the mailbox list

Write-Host -ForegroundColor White "Collecting mailbox list"

if($all) { $mailboxes = @(Get-Mailbox -resultsize unlimited -IgnoreDefaultScope) }

if($server) { $mailboxes = @(Get-Mailbox -server $server -resultsize unlimited -IgnoreDefaultScope) }

if($database){ $mailboxes = @(Get-Mailbox -database $database -resultsize unlimited -IgnoreDefaultScope) }

if($file) {	$mailboxes = @(Get-Content $file | Get-Mailbox -resultsize unlimited) }

if($mailbox) { $mailboxes = @(Get-Mailbox $mailbox) }

#Get the report

Write-Host -ForegroundColor White "Collecting report data"

$mailboxcount = $mailboxes.count
$i = 0

#Loop through mailbox list and find the aged mailboxes
foreach ($mb in $mailboxes)
{
	$i = $i + 1
	$pct = $i/$mailboxcount * 100
	Write-Progress -Activity "Collecting mailbox details" -Status "Processing mailbox $i of $mailboxcount - $mb" -PercentComplete $pct

	$stats = $mb | Get-MailboxStatistics | Select-Object TotalItemSize,TotalDeletedItemSize,ItemCount,LastLogonTime,LastLoggedOnUserAccount
	$lastlogon = $stats.LastLogonTime

	#This is an aged mailbox, so we want some extra details about the account
	
	$user = Get-User $mb
	$aduser = Get-ADUser $mb.samaccountname -Properties Enabled,AccountExpirationDate

	#Create a custom PS object to aggregate the data we're interested in
	
	$userObj = New-Object PSObject
	$userObj | Add-Member NoteProperty -Name "DisplayName" -Value $mb.DisplayName
	$userObj | Add-Member NoteProperty -Name "Title" -Value $user.Title
	$userObj | Add-Member NoteProperty -Name "Department" -Value $user.Department
	$userObj | Add-Member NoteProperty -Name "Office" -Value $user.Office
	$userObj | Add-Member NoteProperty -Name "Enabled" -Value $aduser.Enabled
	$userObj | Add-Member NoteProperty -Name "Expires" -Value $aduser.AccountExpirationDate
	$userObj | Add-Member NoteProperty -Name "Last Mailbox Logon" -Value $lastlogon
	$userObj | Add-Member NoteProperty -Name "Last Logon By" -Value $stats.LastLoggedOnUserAccount
	$userObj | Add-Member NoteProperty -Name "Item Size (Mb)" -Value $stats.TotalItemSize.Value.ToMB()
	$userObj | Add-Member NoteProperty -Name "Deleted Item Size (Mb)" -Value $stats.TotalDeletedItemSize.Value.ToMB()
	$userObj | Add-Member NoteProperty -Name "Items" -Value $stats.ItemCount
	$userObj | Add-Member NoteProperty -Name "Type" -Value $mb.RecipientTypeDetails
	$userObj | Add-Member NoteProperty -Name "Server" -Value $mb.ServerName
	$userObj | Add-Member NoteProperty -Name "Database" -Value $mb.Database

	
	#Add the object to the report
	$report = $report += $userObj
}

#Catch zero item results
$reportcount = $report.count

if ($reportcount -eq 0)
{
	Write-Host -ForegroundColor Yellow "No mailboxes were found matching that criteria."
}
else
{
	#Output single mailbox report to console, otherwise output to CSV file
	if ($mailbox) 
	{
		$report | Format-List
	}
	else
	{
		$report | Export-Csv -Path $reportfile -NoTypeInformation
		Write-Host -ForegroundColor White "Report written to $reportfile in current path."
		Get-Item $reportfile
	}
}
