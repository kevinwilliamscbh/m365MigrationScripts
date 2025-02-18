
<#********************************************************************
Export Teams permissions,
store export in Blob Container

Usage: TeamsPermissions -SharedAccessToken 0000-0000-0000-0000-000 -BaseURI "https:/storage/container" -ImportFile "https:/storage/container/ImportFile.csv?SAS" | "C:/localfile"

-SharedAccessToken (required)
-BaseURI (optional; if not supplied, will use hardcoded value)
-ImportFile (optional; if supplied, will override $teamMailNickName variable)
   ImportFile should contain a single MailNickName per line, with no header
   Include SA key in filename if required

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
    write-host "Detected local write"
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

#Check if NickNames passed in arguments or if need to read
If ($MailNickNames -ne "")
    {
    $MailNickNames = $MailNickNames.Replace(" ","")
    $MailNickNames = $MailNickNames.Replace(","," ")
    $teamMailNickNames = [array]$MailNickNames.Split(" ")
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
            Connect-MicrosoftTeams 
            }
            else
            {
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
