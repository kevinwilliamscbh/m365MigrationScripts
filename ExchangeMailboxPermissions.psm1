<#
.SYNOPSIS
	This is a PowerShell module for Cherry Bekaert.
	cbh.com
	msdn.microsoft.com/powershell

.DESCRIPTION
    The script will export Exchange Shared Mailbox Permissions

.NOTES
    Version:        1.0.1
    Author:         Kevin Williams
    Website:        cbh.com
    Creation Date:  2/15/2025
    Purpose/Change: Initial script development

    Update Date: 2/18/2025
    Purpose/Change: Continued improvements and bug fixes

.CHECKSUM
    Use $Script.ToString().GetHashCode() to verify checksum

.USAGE
	
--->
	##Pass arguments to Invoke-Command
	$ClientCode = 'CBH'
	$SAS = ''
	$ImportFile = ''
	$ExportFile = ''
	$BaseUri = ''
	$Mailboxes = ''
    $RecipientTypeDetails = ''
	$Arguments = @($ClientCode, $SAS, $ImportFile, $ExportFile, $BaseUri, $Mailboxes, $RecipientTypeDetails)
	##Invoke PowerShell Script 
	$URI = 'https://raw.githubusercontent.com/kevinwilliamscbh/m365MigrationScripts/refs/heads/main/ExchangeMailboxPermissions.psm1'
	$Script = [ScriptBlock]::Create((new-object Net.WebClient).DownloadString($URI))	
	Invoke-command -ScriptBlock $Script -ArgumentList $Arguments
--->

	How to use:
	Fill in required arguments for script, then select entire script block and paste into PowerShell.
	Can paste into PowerShell Desktop or PowerShell Core (Azure Cloud CLI).
	
	Arguments:
	ClientCode:  Prepended to export file name if not provided in ExportFile **REQUIRED**
	SAS: Shared Access Key and is required if not provided with ImportFile or ExportFile (depending on usage)
	ImportFile:  CSV file containing objects to import. Can be 'importfile.csv', 'c:\data\importfile.csv', 
                'https://storage/importfile.csv?sas', or 'https://storage/importfile.csv' (depending on usage)
	ExportFile: CSV file contained exported data. Can be 'exportfile.csv', 'c:\data\exportfile.csv', 
                'https://storage/exportfile.csv?sas', or 'https://storage/exportfile.csv' (depending on usage)
	BaseUri: BASE URI used to construct URIs if full paths not provided.
	Mailboxes: Contains string of Shared Mailbox UPNs for export of permissions. Ex: 'FinanceTeam@cbh.com,Project@cbh.com,Demo@cbh.com'
    RecpipientTypeDetails: The allowed values are: 'RoomMailbox', 'EquipmentMailbox', 'SchedulingMailbox', 'LegacyMailbox', 
            'LinkedMailbox', 'LinkedRoomMailbox', 'UserMailbox', 'DiscoveryMailbox', 'TeamMailbox', 'SharedMailbox', 'GroupMailbox', or '' (for all)
	
	Examples:
	
	Example 1 - Retrieve all Shared Mailbox permissions, store export file locally.
	*Desktop will store file in current directory.
	*Azure Cloud Shell will store file in %HOME% directory.
	
	$ClientCode = 'CBH'
	$SAS = ''
	$ImportFile = ''
	$ExportFile = ''
	$BaseUri = ''
	$Mailboxes = ''
    $RecipientTypeDetails = ''
	
	Example 2 - Retrieve specified shared mailbox permissions, store export file locally.
	*Desktop will store file in current directory.
	*Azure Cloud Shell will store file in %HOME% directory.
	
	$ClientCode = 'CBH'
	$SAS = ''
	$ImportFile = ''
	$ExportFile = ''
	$BaseUri = ''
	$Mailboxes = 'FinanceTeam@cbh.com,Project@cbh.com,Demo@cbh.com'
	
	Example 3 - Retrieve specified shared mailbox permissions, store in specified Azure Storage account
	using URI constructed from BaseUri, generated filename, and SAS.
	
	$SAS = 'sp=shareaccesskey1fromstorageaccount'
	$ImportFile = ''
	$ExportFile = ''
	$BaseUri = 'https://storage.blob.core.windows.net/uploaddata'
	$Mailboxes = 'FinanceTeam@cbh.com,Project@cbh.com,Demo@cbh.com'
	
	Example 4 - Retrieve specified shared mailbox permissions, store in specified Azure Storage account
	using URI constructed from BaseUri, provided filename, and SAS.
	
	$SAS = 'sp=shareaccesskey1fromstorageaccount'
	$ImportFile = ''
	$ExportFile = 'myFileName.csv'
	$BaseUri = 'https://storage.blob.core.windows.net/uploaddata'
	$Mailboxes = 'FinanceTeam@cbh.com,Project@cbh.com,Demo@cbh.com'
	
	Example 5 - Retrieve specified shared mailbox permissions, store in specified Azure Storage account
	using URI provided by ExportFile.
	
	$SAS = 'sp=shareaccesskey1fromstorageaccount'
	$ImportFile = ''
	$ExportFile = 'https://storage.blob.core.windows.net/uploaddata/myExportFile.csv?sp=storageaccesskey'
	$BaseUri = ''
	$Mailboxes = 'FinanceTeam@cbh.com,Project@cbh.com,Demo@cbh.com'

    More Examples to Come!
	
#>	

param (
  [Parameter(Mandatory)][String]$ClientCode,
  [string]$SharedAccessToken, 
  [string]$ImportFile,
  [string]$ExportFile,
  [string]$BaseUri,
  [string]$Mailboxes,
  [string]$RecipientTypeDetails
)
$timeStamp = (Get-Date).ToString("yyMMdd_HHmm")
$ErrorActionPreference = "Stop"
$exportData = "Mailbox,User,Permission`n"
$exchangeMailboxes = $null
$readMailboxes = $false
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
            $exportFileUri = $ExportFile
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

    If ($Mailboxes.Contains(" "))
        {
        $exchangeMailboxes = [array]$Mailboxes.split(" ")
        }
    else
        {
        $exchangeMailboxes = $Mailboxes
        }
    }
else
    {
    $readMailboxes = $true
    }

#Import Shared Mailbox list, if provided
If ($ImportFile -ne "")
    {
    $readMailboxes = $false
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
If ($readMailboxes)
    {
    If ($RecipientTypeDetails.Length -eq 0)
        {
        Write-Host "`nReading All mailboxes in tenant" -ForegroundColor Yellow
        $exchangeMailboxes = (Get-Mailbox -ResultSize unlimited).UserPrincipalName | ?{$_ -notlike "DiscoverySearchMailbox*"} | Sort
        }
    else
        {
        Write-Host "`nReading all $($RecipientTypeDetails)es in tenant" -ForegroundColor Yellow
        $exchangeMailboxes = (Get-Mailbox -RecipientTypeDetails $RecipientTypeDetails -ResultSize unlimited).UserPrincipalName | ?{$_ -notlike "DiscoverySearchMailbox*"} | Sort
        }
    }

#Begin collecting shared mailbox statistics
ForEach ($mailbox in $exchangeMailboxes)
    {
    $mailbox = $mailbox.Replace(" ","")
    Write-Host "`nProcessing mailbox $mailbox" -ForegroundColor Yellow
    try
        {
        $permissions = Get-MailboxPermission -Identity $mailbox -ErrorAction SilentlyContinue | ?{$_.User -ne "NT AUTHORITY\SELF"}
        $sendAsPerms = Get-RecipientPermission -Identity $mailbox -ErrorAction SilentlyContinue | ?{$_.Trustee -ne "NT AUTHORITY\SELF"}
        }
        catch
            {
            Write-host "$mailbox not found" -ForegroundColor Red
            }
    If ($permissions -ne $null)
        {
        ForEach ($user in $permissions)
            {
            ForEach ($permission in ([array]$user.AccessRights.split(",").trim()))
                {
                $displayLine = $mailbox + "," + $user.user + "," + $permission
                $outline =  $displayLine + "`n"
                $exportData += $outline
                write-host $displayLine
                }
            }
        }

    If ($sendAsPerms -ne $null)
        {
        ForEach ($SendAs in $sendAsPerms)
            {
            ForEach ($access in ([array]$SendAs.AccessRights.split(",").trim()))
                {
                $displayLine = $mailbox + "," + $SendAs.Trustee + "," + $access
                $outline =  $displayLine + "`n"
                $exportData += $outline
                write-host $displayLine
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

