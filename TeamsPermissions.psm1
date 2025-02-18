

$URI = 'https://raw.githubusercontent.com/kevinwilliamscbh/m365MigrationScripts/refs/heads/main/TeamsPermissions.psm1'
$Script = [ScriptBlock]::Create((new-object Net.WebClient).DownloadString($URI))
$ClientCode = 'CBH'

$SAS = 'sp=racw&st=2025-02-17T18:05:17Z&se=2025-02-22T02:05:17Z&spr=https&sv=2022-11-02&sr=c&sig=ZA6VqImomwOjpwoSSkoLxVZDphD5GJsrNA8ETJisrpI%3D'
$ImportFile = 'https://raw.githubusercontent.com/kevinwilliamscbh/m365MigrationScripts/refs/heads/main/MyTeamsImportFile.csv'
$ExportFile = ''
$BaseUri = 'https://bittitanmigrationangeion.blob.core.windows.net/uploaddata'
#$MailNickNames = 'ProjectDemo12,DemoProject4,ScottMaddenTestDemo,ComplexProcessManagementDemo,ManagedServices,SMProjectTemplate'
$MailNickNames = '' 
                    
$Arguments = @($SAS, $ClientCode, $ImportFile, $ExportFile, $BaseUri, $MailNickNames)
Invoke-command -ScriptBlock $Script -ArgumentList $Arguments


$ClientCode = 'CBH'
$SAS = ''
$ImportFile = ''
$ExportFile = ''
$BaseUri = ''
$MailNickNames = 'ProjectDemo12'
$Arguments = @($ClientCode, $SAS, $ImportFile, $ExportFile, $BaseUri, $MailNickNames)
#cd 'c:\Users\Kevin.Williams\OneDrive - Cherry Bekaert\Scripts\MigrationScripts'
.\TeamsPermissions.ps1 @Arguments
