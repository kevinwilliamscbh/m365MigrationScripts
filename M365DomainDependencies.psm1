#------------------------------------------------------------------------------


# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT

# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT

# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS

# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR

# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.


#------------------------------------------------------------------------------

#

# PowerShell Source Code

#

# NAME:

#    Domain Association Search.PS1

#

# VERSION:

#    1.0

#

#------------------------------------------------------------------------------


#----------------------------------------------

#region Import Assemblies

#----------------------------------------------

[void][Reflection.Assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

[void][Reflection.Assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

[void][Reflection.Assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')

[void][Reflection.Assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')

#endregion Import Assemblies


#Define a Param block to use custom parameters in the project

#Param ($CustomParameter)


function Main {


       Param ([String]$Commandline)


       if((Show-MainForm_psf) -eq 'OK')

       {


       }


       $script:ExitCode = 0 #Set the exit code for the Packager

}



function Show-MainForm_psf

{


       #----------------------------------------------

       #region Import the Assemblies

       #----------------------------------------------

       [void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

       [void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

       [void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')

       #endregion Import Assemblies


       #----------------------------------------------

       #region Generated Form Objects

       #----------------------------------------------

       [System.Windows.Forms.Application]::EnableVisualStyles()

       $form1 = New-Object 'System.Windows.Forms.Form'

       $textbox1 = New-Object 'System.Windows.Forms.TextBox'

       $listbox3 = New-Object 'System.Windows.Forms.ListBox'

       $listbox2 = New-Object 'System.Windows.Forms.ListBox'

       $listbox1 = New-Object 'System.Windows.Forms.ListBox'

       $labelAZUREApps = New-Object 'System.Windows.Forms.Label'

       $labelEXOMailboxesMailUser = New-Object 'System.Windows.Forms.Label'

       $labelMSODSUsersGroupsCont = New-Object 'System.Windows.Forms.Label'

       $labelDomainAssociationSea = New-Object 'System.Windows.Forms.Label'

       $labelDomainName = New-Object 'System.Windows.Forms.Label'

       $buttonSearch = New-Object 'System.Windows.Forms.Button'

       $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

       #endregion Generated Form Objects



       $form1_Load = {

    #Connections

       $Module = Get-Module


    $ModTMP="| "

    foreach ($Mod in $Module){$ModTMP += $Mod.Name +" | "}


    if($ModTMP.IndexOf("| MSOnline |") -gt -1) {

    Import-Module Msonline

    Write-Host "------------- Import MSOnline -------------" -fore cyan

    }

    else {

    Install-Module msonline -Force

    Import-Module Msonline

    Write-Host "------------- Install MSOnline -------------" -fore cyan

    }


    if($ModTMP.IndexOf("| AzureAD |") -gt -1)  {

    Import-Module azuread

     Write-Host "------------- Import AzureAD -------------" -fore cyan


    }

    else {

    Install-Module azuread -Force

    Import-Module azuread

    Write-Host "------------- Install AzureAD -------------" -fore cyan

    }


             $livecred = Get-Credential

             Connect-MsolService -Credential $livecred

             $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powerShell-liveID?serializationLevel=Full -Credential $LiveCred -Authentication Basic –AllowRedirection

             Import-PSSession $Session -AllowClobber

             Connect-AzureAD -Credential $livecred

       }



       $labelDomainAssociationSea_Click={


       }


       $labelDomainName_Click={


       }


       $textbox1_TextChanged={


       }


       $buttonSearch_Click = {

             #Clear the Boxes

             $listbox1.Items.Clear()

             $listbox2.Items.Clear()

             $listbox3.Items.Clear()


             $form1.Cursor = 'WaitCursor'


             #Run MSODS

             $listbox1.Items.Add("Finding active Users...")

             $msods = Get-Msoluser -all | Where-Object { $_.UserPrincipalName -match $textbox1.text }

             $msods2 = Get-MsolUser -all | Where-Object { $_.ProxyAddresses -match $textbox1.text }

             $msods3 = Get-MsolUser -all | Where-Object { $_.WindowsLiveId -match $textbox1.text }

             if ($msods -ne $null) { $listbox1.Items.AddRange($msods.Displayname) }

             if ($msods2 -ne $null) { $listbox1.Items.AddRange($msods2.Displayname) }

             if ($msods3 -ne $null) { $listbox1.Items.AddRange($msods3.Displayname) }


             $listbox1.Items.Add("Finding deleted Users...")

             $msods6 = Get-Msoluser -returndeletedusers -all | Where-Object { $_.UserPrincipalName -match $textbox1.text }

             $msods8 = Get-MsolUser -returndeletedusers -all | Where-Object { $_.ProxyAddresses -match $textbox1.text }

             $msods9 = Get-MsolUser -returndeletedusers -all | Where-Object { $_.WindowsLiveId -match $textbox1.text }

             if ($msods6 -ne $null) { $listbox1.Items.AddRange($msods6.Displayname) }

             if ($msods8 -ne $null) { $listbox1.Items.AddRange($msods8.Displayname) }

             if ($msods9 -ne $null) { $listbox1.Items.AddRange($msods9.Displayname) }


             $listbox1.Items.Add("Finding Groups...")

             $msods11 = Get-MsolGroup -all | Where-Object { $_.EmailAddresses -match $textbox1.text }

             if ($msods11 -ne $null) { $listbox1.Items.AddRange($msods11.Displayname) }


             $listbox1.Items.Add("Finding Contacts...")

             $msods4 = Get-MsolContact -all | Where-Object { $_.UserPrincipalName -match $textbox1.text }

             if ($msods4 -ne $null) { $listbox1.Items.AddRange($msods4.Displayname) }


             #Run EXO

             $listbox2.Items.Add("Finding MAilboxes...")

             $exo15 = Get-Mailbox -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             $exo17 = Get-Mailbox -Softdeleted -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             $exo19 = Get-Mailbox -InactiveMailboxOnly -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             $exo25 = Get-Sitemailbox -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             if ($exo15 -ne $null) { $listbox2.Items.AddRange($exo15.Displayname) }

             if ($exo17 -ne $null) { $listbox2.Items.AddRange($exo17.Displayname) }

             if ($exo19 -ne $null) { $listbox2.Items.AddRange($exo19.Displayname) }

             if ($exo25 -ne $null) { $listbox2.Items.AddRange($exo25.Displayname) }


             $listbox2.Items.Add("Finding Mail Users...")

             $exo11 = Get-MailUser -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             if ($exo11 -ne $null) { $listbox2.Items.AddRange($exo11.Displayname) }


             $listbox2.Items.Add("Finding Groups...")

             $exo1 = Get-DistributionGroup -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             $exo3 = Get-DynamicDistributionGroup -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             $exo5 = Get-UnifiedGroup -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             if ($exo1 -ne $null) { $listbox2.Items.AddRange($exo1.Displayname) }

             if ($exo3 -ne $null) { $listbox2.Items.AddRange($exo3.Displayname) }

             if ($exo5 -ne $null) { $listbox2.Items.AddRange($exo5.Displayname) }


             $listbox2.Items.Add("Finding Groups...")

             $exo9 = Get-MailContact -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             if ($exo9 -ne $null) { $listbox2.Items.AddRange($exo9.Displayname) }


             $listbox2.Items.Add("Finding Public Folders...")

             $exo21 = Get-Mailbox -PublicFolder -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             $exo23 = Get-MailPublicFolder -ResultSize unlimited | Where-Object { $_.EmailAddresses -match $textbox1.text }

             if ($exo21 -ne $null) { $listbox2.Items.AddRange($exo21.Displayname) }

             if ($exo23 -ne $null) { $listbox2.Items.AddRange($exo23.Displayname) }


             #Run Azure

             $listbox3.Items.Add("Finding Applications...")

             $azure = Get-AzureADApplication | Where-Object { $_.IdentifierUris -match $textbox1.text }

             if ($azure -ne $null) { $listbox3.Items.AddRange($azure.Displayname) }



             $form1.Cursor = 'Default'


       }



       #region Control Helper Functions

       function Update-ListBox

       {


             param

             (

                    [Parameter(Mandatory = $true)]

                    [ValidateNotNull()]

                    [System.Windows.Forms.ListBox]

                    $ListBox,

                    [Parameter(Mandatory = $true)]

                    [ValidateNotNull()]

                    $Items,

                    [Parameter(Mandatory = $false)]

                    [string]

                    $DisplayMember,

                    [switch]

                    $Append

             )


             if (-not $Append)

             {

                    $listBox.Items.Clear()

             }


             if ($Items -is [System.Windows.Forms.ListBox+ObjectCollection] -or $Items -is [System.Collections.ICollection])

             {

                    $listBox.Items.AddRange($Items)

             }

             elseif ($Items -is [System.Collections.IEnumerable])

             {

                    $listBox.BeginUpdate()

                    foreach ($obj in $Items)

                    {

                          $listBox.Items.Add($obj)

                    }

                    $listBox.EndUpdate()

             }

             else

             {

                    $listBox.Items.Add($Items)

             }


             $listBox.DisplayMember = $DisplayMember

       }

       #endregion


       $form1_FormClosing=[System.Windows.Forms.FormClosingEventHandler]{

       #Event Argument: $_ = [System.Windows.Forms.FormClosingEventArgs


       }



       #----------------------------------------------

       #region Generated Events

       #----------------------------------------------


       $Form_StateCorrection_Load=

       {

             #Correct the initial state of the form to prevent the .Net maximized form issue

             $form1.WindowState = $InitialFormWindowState

       }


       $Form_StoreValues_Closing=

       {

             #Store the control values

             $script:MainForm_textbox1 = $textbox1.Text

             $script:MainForm_listbox3 = $listbox3.SelectedItems

             $script:MainForm_listbox2 = $listbox2.SelectedItems

             $script:MainForm_listbox1 = $listbox1.SelectedItems

       }



       $Form_Cleanup_FormClosed=

       {

             #Remove all event handlers from the controls

             try

             {

                    $textbox1.remove_TextChanged($textbox1_TextChanged)

                   $labelDomainAssociationSea.remove_Click($labelDomainAssociationSea_Click)

                    $labelDomainName.remove_Click($labelDomainName_Click)

                    $buttonSearch.remove_Click($buttonSearch_Click)

                    $form1.remove_FormClosing($form1_FormClosing)

                    $form1.remove_Load($form1_Load)

                    $form1.remove_Load($Form_StateCorrection_Load)

                    $form1.remove_Closing($Form_StoreValues_Closing)

                    $form1.remove_FormClosed($Form_Cleanup_FormClosed)

             }

             catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }

       }

       #endregion Generated Events



       #----------------------------------------------

       #region Generated Form Code

       #----------------------------------------------

       $form1.SuspendLayout()

       #

       # form1

       #

       $form1.Controls.Add($textbox1)

       $form1.Controls.Add($listbox3)

       $form1.Controls.Add($listbox2)

       $form1.Controls.Add($listbox1)

       $form1.Controls.Add($labelAZUREApps)

       $form1.Controls.Add($labelEXOMailboxesMailUser)

       $form1.Controls.Add($labelMSODSUsersGroupsCont)

       $form1.Controls.Add($labelDomainAssociationSea)

       $form1.Controls.Add($labelDomainName)

       $form1.Controls.Add($buttonSearch)

       $form1.AutoScaleDimensions = '11, 20'

       $form1.AutoScaleMode = 'Font'

       $form1.ClientSize = '1075, 549'

       $form1.Font = 'Microsoft Sans Serif, 10.2pt, style=Bold'

       $form1.Margin = '6, 5, 6, 5'

       $form1.Name = 'form1'

       $form1.Text = 'Form'

       $form1.add_FormClosing($form1_FormClosing)

       $form1.add_Load($form1_Load)

       #

       # textbox1

       #

       $textbox1.Font = 'Microsoft Sans Serif, 10.2pt'

       $textbox1.Location = '349, 81'

       $textbox1.Margin = '6, 5, 6, 5'

       $textbox1.Name = 'textbox1'

       $textbox1.Size = '358, 27'

       $textbox1.TabIndex = 13

       $textbox1.add_TextChanged($textbox1_TextChanged)

       #

       # listbox3

       #

       $listbox3.Font = 'Microsoft Sans Serif, 10.2pt'

       $listbox3.FormattingEnabled = $True

       $listbox3.ItemHeight = 20

       $listbox3.Location = '720, 254'

       $listbox3.Margin = '6, 5, 6, 5'

       $listbox3.Name = 'listbox3'

       $listbox3.Size = '312, 264'

       $listbox3.TabIndex = 12

       #

       # listbox2

       #

       $listbox2.Font = 'Microsoft Sans Serif, 10.2pt'

       $listbox2.FormattingEnabled = $True

       $listbox2.ItemHeight = 20

       $listbox2.Location = '381, 254'

       $listbox2.Margin = '6, 5, 6, 5'

       $listbox2.Name = 'listbox2'

       $listbox2.Size = '300, 264'

       $listbox2.TabIndex = 11

       #

       # listbox1

       #

       $listbox1.Font = 'Microsoft Sans Serif, 10.2pt'

       $listbox1.FormattingEnabled = $True

       $listbox1.ItemHeight = 20

       $listbox1.Location = '33, 254'

       $listbox1.Margin = '6, 5, 6, 5'

       $listbox1.Name = 'listbox1'

       $listbox1.Size = '303, 264'

       $listbox1.TabIndex = 10

       #

       # labelAZUREApps

       #

       $labelAZUREApps.AutoSize = $True

       $labelAZUREApps.Font = 'Microsoft Sans Serif, 7.8pt, style=Bold'

       $labelAZUREApps.Location = '730, 206'

       $labelAZUREApps.Margin = '6, 0, 6, 0'

       $labelAZUREApps.Name = 'labelAZUREApps'

       $labelAZUREApps.Size = '58, 32'

       $labelAZUREApps.TabIndex = 9

       $labelAZUREApps.Text = 'AZURE

(Apps)'

       #

       # labelEXOMailboxesMailUser

       #

       $labelEXOMailboxesMailUser.AutoSize = $True

       $labelEXOMailboxesMailUser.Font = 'Microsoft Sans Serif, 7.8pt, style=Bold'

       $labelEXOMailboxesMailUser.Location = '382, 190'

       $labelEXOMailboxesMailUser.Margin = '6, 0, 6, 0'

       $labelEXOMailboxesMailUser.Name = 'labelEXOMailboxesMailUser'

       $labelEXOMailboxesMailUser.Size = '227, 48'

       $labelEXOMailboxesMailUser.TabIndex = 8

       $labelEXOMailboxesMailUser.Text = 'EXO

(Mailboxes, MailUsers, Groups,

Public Folders)'

       #

       # labelMSODSUsersGroupsCont

       #

       $labelMSODSUsersGroupsCont.AutoSize = $True

       $labelMSODSUsersGroupsCont.Font = 'Microsoft Sans Serif, 7.8pt, style=Bold'

       $labelMSODSUsersGroupsCont.Location = '33, 206'

       $labelMSODSUsersGroupsCont.Margin = '6, 0, 6, 0'

       $labelMSODSUsersGroupsCont.Name = 'labelMSODSUsersGroupsCont'

       $labelMSODSUsersGroupsCont.Size = '184, 32'

       $labelMSODSUsersGroupsCont.TabIndex = 7

       $labelMSODSUsersGroupsCont.Text = 'MSODS

(Users, Groups, Contacts)'

       #

       # labelDomainAssociationSea

       #

       $labelDomainAssociationSea.AutoSize = $True

       $labelDomainAssociationSea.Font = 'Microsoft Sans Serif, 12pt, style=Bold'

       $labelDomainAssociationSea.Location = '382, 9'

       $labelDomainAssociationSea.Margin = '6, 0, 6, 0'

       $labelDomainAssociationSea.Name = 'labelDomainAssociationSea'

       $labelDomainAssociationSea.Size = '278, 25'

       $labelDomainAssociationSea.TabIndex = 3

       $labelDomainAssociationSea.Text = 'Domain Association Search'

       $labelDomainAssociationSea.add_Click($labelDomainAssociationSea_Click)

       #

       # labelDomainName

       #

       $labelDomainName.AutoSize = $True

       $labelDomainName.Location = '209, 84'

       $labelDomainName.Margin = '6, 0, 6, 0'

       $labelDomainName.Name = 'labelDomainName'

       $labelDomainName.Size = '127, 20'

       $labelDomainName.TabIndex = 2

       $labelDomainName.Text = 'Domain Name'

       $labelDomainName.add_Click($labelDomainName_Click)

       #

       # buttonSearch

       #

       $buttonSearch.Font = 'Microsoft Sans Serif, 7.8pt, style=Bold'

       $buttonSearch.Location = '730, 78'

       $buttonSearch.Margin = '6, 5, 6, 5'

       $buttonSearch.Name = 'buttonSearch'

       $buttonSearch.Size = '126, 35'

       $buttonSearch.TabIndex = 1

       $buttonSearch.Text = 'Search'

       $buttonSearch.UseVisualStyleBackColor = $True

       $buttonSearch.add_Click($buttonSearch_Click)

       $form1.ResumeLayout()

       #endregion Generated Form Code


       #----------------------------------------------


       #Save the initial state of the form

       $InitialFormWindowState = $form1.WindowState

       #Init the OnLoad event to correct the initial state of the form

       $form1.add_Load($Form_StateCorrection_Load)

       #Clean up the control events

       $form1.add_FormClosed($Form_Cleanup_FormClosed)

       #Store the control values when form is closing

       $form1.add_Closing($Form_StoreValues_Closing)

       #Show the Form

       return $form1.ShowDialog()


}


#Start the application

Main ($CommandLine)


#------------------------------------------------------------------------------