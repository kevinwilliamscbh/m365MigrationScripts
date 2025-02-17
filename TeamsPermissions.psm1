
<#********************************************************************
Export Teams permissions,
store export in Blob Container

Usage: TeamsPermissions -SharedAccessToken 0000-0000-0000-0000-000 -BaseURI "https:/storage/container" -ImportFile "https:/storage/container/ImportFile.csv?SAS" | "C:/localfile"

-SharedAccessToken (required)
-BaseURI (optional; if not supplied, will use hardcoded value)
-ImportFile (optional; if supplied, will override $teamMailNickName variable)
   ImportFile should contain a single MailNickName per line, with no header
   Include SA key in filename if required

*********************************************************************#>

#Require Shared Access Token
param (
  [Parameter(Mandatory)][string]$SharedAccessToken,
  [string]$ClientCode = "CBH",
  [string]$ImportFile,
  [string]$BaseUri,
  [string]$MailNickNames
)

$ErrorActionPreference = "Stop"
$exportFile = "Team,MailNick,Channel,UPN,DisplayName,Role`n"
$teams = @()
$readTeams = $false
$timeStamp = (Get-Date).ToString("yyMMdd_HHmm")
$exportFileName = $ClientCode + "TeamsPermissions-" + $timeStamp + ".csv"
$exportFileUri = "$BaseUri/$exportFileName" + "?" + $SharedAccessToken

#Check if NickNames passed in arguments or if need to read
Write-Host $MailNickNames
Write-host $MailNickNames.Length
Write-host $MailNickNames.ToString()
Write-Host "Split" $MailNickNames.Split(" ") -ForegroundColor Yellow
If ($MailNickNames -ne $null)
    {
    Write-host "Converting argument to array"
    $teamMailNickNames = $MailNickNames.Split(" ")
    Write-Host $teamMailNickNames
    pause
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
    Get-Team
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
    Write-Host "Reading all Teams from tenant"
    $teamMailNickNames = (Get-Team).MailNickName
    }

#Obtain Team Channels
Write-host $teamMailNickNames
ForEach($team in $teamMailNickNames)
    {
    $addTeam = $true
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
        Write-Host "Processing Channel $channel in Team" $team.DisplayName
        ForEach($user in $users)
            {
            $exportFile += $team.DisplayName + "," + $team.MailNick + "," + $channel + "," + $user.User + "," + $user.Name + "," + $user.Role + "`n"
            }
        }
    }

#Upload file to Angeion File Share and verify
$headers = @{'x-ms-blob-type' = 'BlockBlob'}
try
    {
    Invoke-RestMethod -Uri $exportFileUri -Method Put -Body $exportFile -Headers $headers
    }
    catch
        {
        Write-Host "Error Message: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "ExportFileURI:" $exportFileUri
        throw "File upload failed"
        }
try
    {
    $check = Invoke-RestMethod -Uri $exportFileUri -Method Get -Headers $headers
    }
    catch
        {
        Write-Host "Unable to verify file upload" -ForegroundColor Red
        Write-Host "Error Message: $($_.Exception.Message)" -ForegroundColor Yellow
        pause
        }
If ($check -eq $exportFile)
    {
    Write-Host "File successfully uploaded" -ForegroundColor Yellow
    Disconnect-MicrosoftTeams -Confirm:$false
    }
