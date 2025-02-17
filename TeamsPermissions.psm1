
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
  [string]$ImportFile,
  [string]$BaseUri = "https://bittitanmigrationangeion.blob.core.windows.net/uploaddata",
  [string]$ExportFileName
)

$ErrorActionPreference = "Stop"
$exportFile = "Team,MailNick,Channel,UPN,DisplayName,Role`n"
$exportFileUri = "$BaseUri/$exportFileName" + "?" + $sharedAccessToken
$teams = @()
$teamMailNickName = @()
$readTeams = $false

If ($ExportFileName.Length -eq 0)
    {
    $timeStamp = (Get-Date).ToString("yyMMdd_HHmm")
    $exportFileName = "drcTeamsPermissions-" + $timeStamp + ".csv"
    $exportFileUri = "$BaseUri/$exportFileName" + "?" + $SharedAccessToken
    }

#<#----- Delete $teamMailNickName variable to return all Teams
#<#----- Supplying ImportFile will override $teamMailNickName
$teamMailNickName = @(
                      "ProjectDemo12"
                      "DemoProject4"
                      "ScottMaddenTestDemo"
                      "ComplexProcessManagementDemo"
                      "ManagedServices"
                      "SMProjectTemplate"
                        )
#<------------------------------------------------------------#>
#<------------------------------------------------------------#>

#Retrieve Team GroupIDs and DisplayNames *Required for retrieving members*
If ($ImportFile.Length -gt 0)
    {
    try
        {
        $teamMailNickName = Import-CSV -Path $ImportFile
        }
        catch
            {
            Write-Host "Error Message: $($_.Exception.Message)" -ForegroundColor Red
            Throw "Error occured during file import"
            }
    }
else
    {
    if($teamMailNickName.Length -eq 0)
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
    $teamMailNickName = (Get-Team).MailNickName
    }
ForEach($team in $teamMailNickName)
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
        Write-Host "ExportFileURI:"$exportFileUri
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
