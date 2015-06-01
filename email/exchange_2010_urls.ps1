    # Script to allow you to set all virtual directories to a common name like mail.company.com

    Start-Transcript

    # Variables

    [string]$EASExtend = "/Microsoft-Server-ActiveSync"
    [string]$PShExtend = "/powershell"
    [string]$OWAExtend = "/OWA"
    [string]$OABExtend = "/OAB"
    [string]$SCPExtend = "/Autodiscover/Autodiscover.xml"
    [string]$EWSExtend = "/EWS/Exchange.asmx"
    [string]$ECPExtend = "/ECP"
    [string]$ConfirmPrompt = "Set this Value? (Y/N)"
    [string]$NoChangeForeground = "white"
    [string]$NoChangeBackground = "red"

    Write-host "This will allow you to set the virtual directories associated with setting up a single SSL certificate to work with Exchange 2010."
    Write-host ""
    [string]$base = Read-host "Base name of virtual directory (e.g. mail.company.com)"
    write-host ""

    # =============================================
    # Validate if a third party trusted certificate is being used
    # because BITS used by OAB downloads won’t use untrusted certificates
    [string]$set = Read-host "Is the certificate being used an internally generated certificate? (Y/N)"
    Write-host ""

    if ($set -eq "Y")    {
        [string]$OABprefix = "http://"
        [boolean]$OABRequireSSL = $false
    }    else    {
        [string]$OABprefix = "https://"
        [boolean]$OABRequireSSL = $true
    }

    # =============================================
    # Build the OAB URL and set the internal Value

    Write-host "Setting OAB Virtual Directories" -foregroundcolor Yellow
    write-host ""

    $OABURL = $OABprefix + $base + $OABExtend

    [array]$OABCurrent = Get-OABVirtualDirectory

    Foreach ($value in $OABcurrent) {
        Write-host "Looking at Server: " $value.server
        Write-host "Current Internal Value: " $value.internalURL
        Write-host "New Internal Value:     " $OABUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y")    {
            Set-OABVirtualDirectory -id $value.identity -InternalURL $OABURL -RequireSSL:$OABRequireSSL
        } else {
            write-host "OAB Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }

        Write-host "Looking at Server: " $value.server
        Write-host "Current External Value: " $value.externalURL
        Write-host "New External Value:     " $OABUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y") {
            Set-OABVirtualDirectory -id $value.identity -ExternalURL $OABURL -RequireSSL:$OABRequireSSL
        } else {
            write-host "OAB Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }
    }

    # ============================================
    # Build the Autodiscover URL and set the SCP Value

    Write-host "Setting Autodiscover Service Connection Point" -foregroundcolor Yellow
    write-host ""

    $SCPURL = "https://" + $base + $SCPExtend

    [array]$SCPCurrent = Get-ClientAccessServer

    Foreach ($value in $SCPCurrent) {
        Write-host "Looking at Server: " $value.name
        Write-host "Current SCP value: " $value.AutoDiscoverServiceInternalUri.absoluteuri
        Write-host "New SCP Value:     " $SCPURL
        [string]$set = Read-host $ConfirmPrompt
        write-host ""
        if ($set -eq "Y")    {
             Set-ClientAccessServer -id $value.identity -AutoDiscoverServiceInternalUri $SCPURL
        }    else {
            write-host "Autodiscover Service Connection Point internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }
    }

    # =============================================
    # Build the EWS URL and set the internal Value

    Write-host "Setting Exchange Web Services Virtual Directories" -foregroundcolor Yellow
    write-host ""

    $EWSURL = "https://" + $base + $EWSExtend

    [array]$EWSCurrent = Get-WebServicesVirtualDirectory

    Foreach ($value in $EWSCurrent) {
        Write-host "Looking at Server: " $value.server
        Write-host "Current Internal Value: " $value.internalURL
        Write-host "New Internal Value:     " $EWSUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y")    {
            Set-WebServicesVirtualDirectory -id $value.identity -InternalURL $EWSURL
         } else {
            write-host "Exchange Web Services Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
         }

        Write-host "Looking at Server: " $value.server
        Write-host "Current External Value: " $value.externalURL
        Write-host "New External Value:     " $EWSUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y")    {
            Set-WebServicesVirtualDirectory -id $value.identity -ExternalURL $EWSURL
        } else {
            write-host "Exchange Web Services Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }
    }

    # =============================================
    # Build the PowerShell URL and set the internal Value

    Write-host "Setting UM Virtual Directories" -foregroundcolor Yellow
    write-host ""

    $PShURL = "http://" + $base + $PShExtend

    [array]$PShCurrent = Get-PowerShellVirtualDirectory

    foreach ($value in $PShCurrent) {
        Write-host "Looking at Server: " $value.server
        Write-host "Current Internal Value: " $value.internalURL
        Write-host "New Internal Value:     " $PShUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y") {
            Set-PowerShellVirtualDirectory -id $value.identity -InternalURL $PShURL
        } else {
            write-host "PowerShell Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }

        Write-host "Looking at Server: " $value.server
        Write-host "Current External Value: " $value.externalURL
        Write-host "New External Value:     " $PShUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y") {
            Set-PowerShellVirtualDirectory -id $value.identity -ExternalURL $PShURL
        } else {
            write-host "PowerShell Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }
    }

    # =============================================
    # Build the ECP URL and set the internal Value

    Write-host "Setting ECP Virtual Directories" -foregroundcolor Yellow
    write-host ""

    $ECPURL = "https://" + $base + $ECPExtend

    [array]$ECPCurrent = Get-ECPVirtualDirectory

    foreach ($value in $ECPCurrent) {
        Write-host "Looking at Server: " $value.server
        Write-host "Current Internal Value: " $value.internalURL
        Write-host "New Internal Value:     " $ECPUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y") {
            Set-ECPVirtualDirectory -id $value.identity -InternalURL $ECPURL
        } else {
            write-host "ECP Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }

        Write-host "Looking at Server: " $value.server
        Write-host "Current External Value: " $value.externalURL
        Write-host "New External Value:     " $ECPUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y") {
            Set-ECPVirtualDirectory -id $value.identity -ExternalURL $ECPURL
        } else {
            write-host "ECP Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }
    }

    # =============================================
    # Build the OWA URL and set the internal Value

    Write-host "Setting OWA Virtual Directories" -foregroundcolor Yellow
    write-host ""

    $OWAURL = "https://" + $base + $OWAExtend

    [array]$OWACurrent = Get-OWAVirtualDirectory

    foreach ($value in $OWACurrent) {
        Write-host "Looking at Server: " $value.server
        Write-host "Current Internal Value: " $value.internalURL
        Write-host "New Internal Value:     " $OWAUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y") {
            Set-OWAVirtualDirectory -id $value.identity -InternalURL $OWAURL
        } else {
            write-host "OWA Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }

        Write-host "Looking at Server: " $value.server
        Write-host "Current External Value: " $value.externalURL
        Write-host "New External Value:     " $OWAUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y") {
            Set-OWAVirtualDirectory -id $value.identity -ExternalURL $OWAURL
        } else {
            write-host "OWA Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }
    }

    # =============================================
    # Build the EAS URL and set the internal Value

    Write-host "Setting EAS Virtual Directories" -foregroundcolor Yellow
    write-host ""

    $EASURL = "https://" + $base + $EASExtend

    [array]$EASCurrent = Get-ActiveSyncVirtualDirectory

    foreach ($value in $EASCurrent) {
        Write-host "Looking at Server: " $value.server
        Write-host "Current Internal Value: " $value.internalURL
        Write-host "New Internal Value:     " $EASUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y") {
            Set-ActiveSyncVirtualDirectory -id $value.identity -InternalURL $EASURL
        } else {
            write-host "EAS Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }

        Write-host "Looking at Server: " $value.server
        Write-host "Current External Value: " $value.externalURL
        Write-host "New External Value:     " $EASUrl
        [string]$set = Read-host $ConfirmPrompt
        write-host ""

        if ($set -eq "Y") {
            Set-ActiveSyncVirtualDirectory -id $value.identity -ExternalURL $EASURL
        } else {
            write-host "EAS Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
        }
    }
    Stop-Transcript
