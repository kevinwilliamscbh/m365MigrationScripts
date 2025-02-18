<#
.SYNOPSIS
	This is a PowerShell module for Cherry Bekaert.
	cbh.com
	msdn.microsoft.com/powershell

.DESCRIPTION
    The script will export Team Channel Permissions

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
	
	##Pass arguments to Invoke-Command
	$ClientCode = 'CBH'
	$SAS = ''
	$ImportFile = ''
	$ExportFile = ''
	$BaseUri = ''
	$MailNickNames = ''
	$Arguments = @($ClientCode, $SAS, $ImportFile, $ExportFile, $BaseUri, $MailNickNames)
	##Invoke PowerShell Script 
	$URI = 'https://raw.githubusercontent.com/kevinwilliamscbh/m365MigrationScripts/refs/heads/main/TeamsPermissions.psm1'
	$Script = [ScriptBlock]::Create((new-object Net.WebClient).DownloadString($URI))	
	Invoke-command -ScriptBlock $Script -ArgumentList $Arguments

	How to use:
	Fill in required arguments for script, then select entire script block and paste into PowerShell.
	Can paste into PowerShell Desktop or PowerShell Core (Azure Cloud CLI).
	
	Arguments:
	ClientCode:  Prepended to export file name if not provided in ExportFile **REQUIRED**
	SAS: Shared Access Key and is required if not provided with ImportFile or ExportFile (depending on usage)
	ImportFile:  CSV file containing objects to import. Can be 'importfile.csv', 'c:\data\importfile.csv', 
            'https://storage/importfile.csv?sas', or 'https://storage/importfile.csv' (depending on usage)
	ExportFile: CSV file contained exported data. Can be 'exportfile.csv', 
            'c:\data\exportfile.csv', 'https://storage/exportfile.csv?sas', or 'https://storage/exportfile.csv' (depending on usage)
	BaseUri: BASE URI used to construct URIs if full paths not provided.
	MailNickNames: Contains string of Team MailNickNames for export of permissions. Ex: 'FinanceTeam,Project,Demo'
	
	Examples:
	
	Example 1 - Retrieve all Team permissions, store export file locally.
	*Desktop will store file in current directory.
	*Azure Cloud Shell will store file in %HOME% directory.
	
	$ClientCode = 'CBH'
	$SAS = ''
	$ImportFile = ''
	$ExportFile = ''
	$BaseUri = ''
	$MailNickNames = ''
	
	Example 2 - Retrieve specified Team permissions, store export file locally.
	*Desktop will store file in current directory.
	*Azure Cloud Shell will store file in %HOME% directory.
	
	$ClientCode = 'CBH'
	$SAS = ''
	$ImportFile = ''
	$ExportFile = ''
	$BaseUri = ''
	$MailNickNames = 'TeamMailNick1,TeamMailNick2, TeamMailNick3'
	
	Example 3 - Retrieve specified Team permissions, store in specified Azure Storage account
	using URI constructed from BaseUri, generated filename, and SAS.
	
	$SAS = 'sp=shareaccesskey1fromstorageaccount'
	$ImportFile = ''
	$ExportFile = ''
	$BaseUri = 'https://storage.blob.core.windows.net/uploaddata'
	$MailNickNames = 'TeamMailNick1,TeamMailNick2, TeamMailNick3'
	
	Example 4 - Retrieve specified Team permissions, store in specified Azure Storage account
	using URI constructed from BaseUri, provided filename, and SAS.
	
	$SAS = 'sp=shareaccesskey1fromstorageaccount'
	$ImportFile = ''
	$ExportFile = 'myFileName.csv'
	$BaseUri = 'https://storage.blob.core.windows.net/uploaddata'
	$MailNickNames = 'TeamMailNick1,TeamMailNick2, TeamMailNick3'
	
	Example 5 - Retrieve specified Team permissions, store in specified Azure Storage account
	using URI provided by ExportFile.
	
	$SAS = 'sp=shareaccesskey1fromstorageaccount'
	$ImportFile = ''
	$ExportFile = 'https://storage.blob.core.windows.net/uploaddata/myExportFile.csv?sp=storageaccesskey'
	$BaseUri = ''
	$MailNickNames = 'TeamMailNick1,TeamMailNick2, TeamMailNick3'
	
	More Examples to Come!
	
#>

param (
  [Parameter(Mandatory)][String]$ClientCode,
  [string]$SharedAccessToken,  
  [string]$ImportFile,
  [string]$ExportFile,
  [string]$BaseUri,
  [string]$MailNickNames
)

$ErrorActionPreference = "Stop"
$exportData = "Team,MailNick,Channel,UPN,DisplayName,Role`n"
$teams = @()
$timeStamp = (Get-Date).ToString("yyMMdd_HHmm")
$exportFileName = $ClientCode + "_TeamsPermissions-" + $timeStamp + ".csv"
$exportFileUri = "$BaseUri/$exportFileName" + "?" + $SharedAccessToken
$ProgressPreference = "SilentlyContinue"
$readTeams = $false
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
                $exportFileUri = "$ExportFile" +"?" + $SharedAccessToken
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
                    $exportFileUri = "$BaseUri$ExportFile" +"?" + $SharedAccessToken
                    }
                else
                    {
                    $exportFileUri = "$BaseUri/$ExportFile" +"?" + $SharedAccessToken
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

#Check if NickNames passed in arguments or if need to read
If ($MailNickNames -ne "")
    {
    $MailNickNames = $MailNickNames.Replace(" ","")
    $MailNickNames = $MailNickNames.Replace(","," ")
    If ($MailNickNames.Contains(" "))
        {
        $teamMailNickNames = [array]$MailNickNames.Split(" ")
        }
     else
        {
        $teamMailNickNames = $MailNickNames
        }
    }
else
    {
    $readTeams = $true
    }

#Check to see if file import required
If ($ImportFile.Length -gt 0)
    {
    If ($ImportFile.Substring(0,5) -eq "https")
        {
        $teamMailNickNames = (Invoke-WebRequest -Uri $ImportFile -Method Get).Content
        }
    else
        {
        $teamMailNickNames = Import-CSV -Path $ImportFile
        }
    }
else
    {
    if($teamMailNickNames.Length -eq 0)
        {
        $readTeams = $true
        }
    }

#Connect to MS Teams PowerShell
Try
    {
    $null = Get-Team
    }
    catch
        {
        If ($PSVersionTable.PSEdition -eq "Desktop")
            {
            #Need to check installed version
            Try
                {
                $eom = Get-InstalledModule -Name "MicrosoftTeams"
                If ([int]::Parse($eom.version.Replace(".","")) -lt 670)
                    {
                    Throw "Please update Microsoft Teams module to Verion 6.7.0"
                    }
                else
                    {
                    Import-Module -Name "MicrosoftTeams" -NoClobber
                    }
                }
                Catch
                    {
                    Throw "Microsoft Teams module not installed"
                    }
            Import-Module MicrosoftTeams
            Connect-MicrosoftTeams 
            }
            else
            {
            #Azure Cloud Shell
            Connect-MicrosoftTeams -UseDeviceAuthentication
            }
        }

If ($readTeams)
    {
    Write-Host "Processing all Teams from tenant" -Foreground Yellow
    $teamMailNickNames = (Get-Team).MailNickName
    }

#Obtain Team Channels
ForEach($team in $teamMailNickNames)
    {
    $addTeam = $true
    $team = $team.replace(" ","")
    try
        {
        $teamInfo = Get-Team -MailNickName $team
        }
        catch
            {
            Write-Host "Error retrieving Team: $($team)" -ForegroundColor Red
            Write-Host "Error Message: $($_.Exception.Message)" -ForegroundColor Red
            $addTeam = $false
            }
    try
        {
        $channels = [array](Get-TeamAllChannel -GroupId $teamInfo.GroupID).DisplayName
        }
        catch
            {
            Write-Host "Error retrieving Channels for: $($team)" -ForegroundColor Red
            Write-Host "Error Message: $($_.Exception.Message)" -ForegroundColor Red
            $addTeam = $false   
            }  
    $teamObj = New-Object PSObject -Property @{
        GroupID     = $teamInfo.GroupID
        DisplayName = $teamInfo.DisplayName
        MailNick    = $team
        Channels    = $channels
        }
    If($addTeam)
        {
        $teams += $teamObj
        }
    }

#Retrieve members and update export data
ForEach($team in $teams)
    {
    ForEach($channel in $team.channels)
        {
        $users = Get-TeamChannelUser -GroupID $team.GroupID -DisplayName $channel
        Write-Host "`nProcessing Channel $channel in Team" $team.DisplayName -ForegroundColor Yellow
        ForEach($user in $users)
            {
            $displayLine = $team.DisplayName + "," + $team.MailNick + "," + $channel + "," + $user.User + "," + $user.Name + "," + $user.Role
            $outLine = $displayLine + "`n"
            $exportData += $outLine
            Write-Host $displayLine
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
        Disconnect-MicrosoftTeams -Confirm:$false
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
        Disconnect-MicrosoftTeams -Confirm:$false
        }
        catch
            {
            Write-Host "`nUnable to write export file" -ForegroundColor Red
            }
    }
