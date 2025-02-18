
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

#Require Shared Access Token
param (
  [Parameter(Mandatory)][String]$ClientCode,
  [string]$SharedAccessToken, 
  [string]$ImportFile,
  [string]$ExportFile,
  [string]$BaseUri,
  [string]$Mailboxes
)
$timeStamp = (Get-Date).ToString("yyMMdd_HHmm")
$ErrorActionPreference = "Stop"
$exportData = "Mailbox,User,Permission`n"
$exchangeMailboxes = $null
$readSharedMailboxes = $false
$exportFileName = $ClientCode + "_MailboxPermissions-" + $timeStamp + ".csv"
$exportFileUri = "$BaseUri/$exportFileName" + "?" + $SharedAccessToken
$writeLocal = $false

#Check if need to write local
If (($SharedAccessToken.Length -eq 0) -and ($ExportFile.Length -eq 0) -and ($BaseUri.Length -eq 0))
    {
    $writeLocal = $true
    }

#Check if need to update ExportFileName
If ($ExportFile.Length -gt 0)
    {
    #export file set
    If ($ExportFile.Substring(0,5) -eq "https")
        {
        #using https
        If ($ExportFile.Contains("?s"))
            {
            #exportfile contains https and sas token, no more work
            $exportFileName = $ExportFile
            }
            else
            {
            #SAS not present
            If ($SharedAccessToken.Length -gt 0)
                {
                #SAS token provided
                #Construct URL of Exportfile + SAS
                #Remove trailing slash if detected
                If ($ExportFile.LastIndexOf("/") -eq ($ExportFile.Length-1))
                    {
                    $ExportFile = $exportFile.Substring(0,($exportFile.Length-1))
                    }
                $exportFileName = "$ExportFile" +"?" + $SharedAccessToken
                }
                else
                {
                #sas token not provided
                Throw "SharedAccessToken missing - unable to construct export file"
                }
            }
        }
        else
        {
        #not using https
        If ($BaseUri.Length -gt 0)
            {
            #Exportfile set using BaseURL
            If ($SharedAccessToken.Length -gt 0)
                {
                #Exportfile, BaseUriSet, and SAS Set
                #Construct URL
                #Remove trailing slash from Base
                If ($BaseUri.LastIndexOf("/") -eq ($BaseUri.Length-1))
                    {
                    $exportFileName = "$BaseUri$ExportFile" +"?" + $SharedAccessToken
                    }
                else
                    {
                    $exportFileName = "$BaseUri/$ExportFile" +"?" + $SharedAccessToken
                    }
                }
                else
                {
                #SAS not set
                Throw "SharedAccessToken missing - unable to construct export file"
                }
            }
            else
            {
            #Export file set and missing BaseURI
            $exportFileName = $ExportFile
            $writeLocal = $true
            #Throw "BASEUuri missing - unable to construct export file"
            }
        }    
    }

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
        write-host "Imported mailboxes via Storage Api"
        }
    else
        {
        $exchangeMailboxes = Import-CSV -Path $ImportFile
        Write-Host "Imported mailboxes via local file"
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
    $mailbox = $mailbox.Replace(" ","")
    $permissions = Get-MailboxPermission -Identity $mailbox -ErrorAction SilentlyContinue | ?{$_.User -ne "NT AUTHORITY\SELF"}
    If ($permissions -ne $null)
        {
        ForEach ($user in $permissions)
            {
            ForEach ($permission in ([array]$user.AccessRights.split(",").trim()))
                {
                $displayLine = $mailbox + "," + $user.user + "," + $permission
                $outline =  $displayLine + "`n"
                $exportData += $outline
                write-host $outline
                }
            }
        }
    }

#Write destination file
If ($writeLocal)
    {
    try
        {
        Out-File -FilePath $exportFileName -Encoding ascii -InputObject $exportData -Force
        Write-Host "`nFile $exportFileName successfully created" -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false
        }
            catch
            {
            Write-Host "`nUnable to write export file" -ForegroundColor Red
            }
    }
else
    {
    try
        {
        $headers = @{'x-ms-blob-type' = 'BlockBlob'}
        Invoke-RestMethod -Uri $exportFileUri -Method Put -Body $exportData -Headers $headers
        Write-Host "`nFile $exportFileName successfully created" -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false
        }
        catch
            {
            Write-Host "`nUnable to write export file" -ForegroundColor Red
            }
    }

