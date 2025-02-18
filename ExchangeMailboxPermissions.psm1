﻿
<#********************************************************************
Export Exchange shared mailbox permissions,
store export in Blob Container

Usage: ExchangeMailboxPermissions -SharedAccessToken 0000-0000-0000-0000-000 -BaseURI "https:/storage/container" -ImportFile "https:/storage/container/ImportFile.csv?SAS" | "C:/localfile"

-SharedAccessToken (required)
-BaseURI (optional; if not supplied, will use hardcoded value)
-ImportFile (optional; if supplied, will override $teamMailNickName variable)
   ImportFile should contain a single Shared Mailbox UPN per line, with no header
   Include SA key in filename if required
 -Mailboxes (options; if not supplied will pull all shared mailboxes)
 -Domain (required in mailbox option used)

Kevin Williams
CBH
Updated 17FEB2025

*********************************************************************#>

param (
  [Parameter(Mandatory)][string]$SharedAccessToken,
  [string]$ClientCode = 'CBH',
  [string]$ImportFile,
  [string]$BaseUri,
  [string]$Mailboxes
)

$ErrorActionPreference = "Stop"
$exportFile = "Mailbox,User,Permission`n"
$exchangeMailboxes = $null
$readSharedMailboxes = $false
$exportFileName = $ClientCode + "_MailboxPermissions-" + $timeStamp + ".csv"
$exportFileUri = "$BaseUri/$exportFileName" + "?" + $SharedAccessToken

#Check if mailboxes passed in arguments or if need to reach from Exhange Online
If ($Mailboxes -ne "")
    {
    $Mailboxes = $Mailboxes.Replace(" ","")
    $Mailboxes = $Mailboxes.Replace(","," ")
    $exchangeMailboxes = [array]$Mailboxes.split(" ")
    }
else
    {
    $readSharedMailboxes = $true
    }

#Import Shared Mailbox list, if provided
If ($ImportFile -ne "")
    {
    $readSharedMailboxes = $false
    If ($ImportFile.Substring(0,5) -eq "https")
        {
        $exchangeMailboxes = (Invoke-WebRequest -Uri $ImportFile -Method Get).Content
        }
    else
        {
        $exchangeMailboxes = Import-CSV -Path $ImportFile
        }
    }


#Connect to Exchange Online PowerShell
$connectionInfo = Get-ConnectionInformation | ?{$_.Name -like "ExchangeOnline*"}
If ($connectionInfo.State -ne "Connected")
    {
    If ($PSVersionTable.PSEdition -eq "Desktop")
        {
        Try
            {
            $eom = Get-InstalledModule -Name "ExchangeOnlineManagement"
            If ([int]::Parse($eom.version.Replace(".","")) -lt 360)
                {
                Throw "Please update Exchange Online Management module to Verion 3.6.0"
                }
            else
                {
                Import-Module -Name "ExchangeOnlineManagement" -NoClobber
                }
            }
            Catch
                {
                Throw "Exchange Online Managment module not installed"
                }

        Import-Module "ExchangeOnlineManagement" -NoClobber
        Connect-ExchangeOnline -ShowBanner:$false
        }
    else
        {
        Connect-ExchangeOnline -ShowBanner:$false -Device
        }
    }
If ($readSharedMailboxes)
    {
    Write-Host "Reading all Shared mailboxes in tenant"
    $exchangeMailboxes = (Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize unlimited).UserPrincipalName
    }

#Begin collecting shared mailbox statistics
ForEach ($mailbox in $exchangeMailboxes)
    {
    $permissions = Get-MailboxPermission -Identity $mailbox -ErrorAction SilentlyContinue | ?{$_.User -ne "NT AUTHORITY\SELF"}
    If ($permissions -ne $null)
        {
        ForEach ($user in $permissions)
            {
            ForEach ($permission in ([array]$user.AccessRights.split(",").trim()))
                {
                Write-host $user.user
                $exportFile += $mailbox + "," + $user.user + "," + $permission + "`n"
                }
            }
        }
    }

#Upload file to Angeion File Share and verify
$headers = @{'x-ms-blob-type' = 'BlockBlob'}
Invoke-RestMethod -Uri $exportFileUri -Method Put -Body $exportFile -Headers $headers
try
    {
    $check = Invoke-RestMethod -Uri $exportFileUri -Method Get -Headers $headers
    }
catch
    {
    Write-Host "Unable to verify file upload" -ForegroundColor Red
    break
    }
If ($check -eq $exportFile)
    {
    Write-Host "File successfully uploaded" -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    }
