<#

.AUTHORS
/u/bwya77
/u/grizzlywinter
/u/briangig

.SYNOPSIS
This is the source code for o365 Adminsitration Center

.DESCRIPTION
The o365 Admin Center is a GUI application that administrators can use to perform some of the most common o365 tasks. The output (error or success) is sent to the textbox which also acts as a input for custom commands. You can also save the output to a file. 

.NOTES
This is built with a GUI and not a stand alone script.

#>

#region Control Helper Functions
function Load-ComboBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ComboBox.

	.DESCRIPTION
		Use this function to dynamically load items into the ComboBox control.

	.PARAMETER  ComboBox
		The ComboBox control you want to add items to.

	.PARAMETER  Items
		The object or objects you wish to load into the ComboBox's Items collection.

	.PARAMETER  DisplayMember
		Indicates the property to display for the items in this control.
	
	.PARAMETER  Append
		Adds the item(s) to the ComboBox without clearing the Items collection.
	
	.EXAMPLE
		Load-ComboBox $combobox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Load-ComboBox $combobox1 "Red" -Append
		Load-ComboBox $combobox1 "White" -Append
		Load-ComboBox $combobox1 "Blue" -Append
	
	.EXAMPLE
		Load-ComboBox $combobox1 (Get-Process) "ProcessName"
#>
	Param (
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.ComboBox]$ComboBox,
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		$Items,
		[Parameter(Mandatory = $false)]
		[string]$DisplayMember,
		[switch]$Append
	)
	
	if (-not $Append)
	{
		$ComboBox.Items.Clear()
	}
	
	if ($Items -is [Object[]])
	{
		$ComboBox.Items.AddRange($Items)
	}
	elseif ($Items -is [Array])
	{
		$ComboBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$ComboBox.Items.Add($obj)
		}
		$ComboBox.EndUpdate()
	}
	else
	{
		$ComboBox.Items.Add($Items)
	}
	
	$ComboBox.DisplayMember = $DisplayMember
}
#endregion

###FORM ITEMS###

	#Form

$FormO365AdministrationCenter_Load = {
	
	#Sets the text for the button
	$ButtonConnectTo365.Text = "Connect to Office 365"
	
	#Sets the text for the button
	$ButtonDisconnect.Text = "Disconnect"
	
	#Sets the text for the button
	$ButtonExportToFile.Text = "Export to File"
	
	#Sets the text for the form
	$FormO365AdministrationCenter.Text = "O365 Administration Center"
	
	#Allows copy/paste
	$TextboxResults.ShortcutsEnabled = $True
	
	#Sets the dialog result
	$ButtonRunCustomCommand.DialogResult = 'None'
	
	#Sets the default button
	$FormO365AdministrationCenter.acceptbutton = $ButtonRunCustomCommand
	
	#Disabled connect to partner button
	$PartnerConnectButton.Enabled = $False
	
	#Disabled disconnect from partner button
	#$ButtonDisconnectFromPartner.Enabled = $False
	
	#Alphabitcally sorts combobox
	$PartnerComboBox.Sorted = $True
	
	#Disables word wrap on the text box
	$TextboxResults.WordWrap = $False
	
	$ButtonDisconnect.Enabled = $False
	
	#Place objects on the bottom
	$ButtonExportToFile.Anchor = 'Bottom'
	$ButtonConnectTo365.Anchor = 'Bottom'
	$Partner_Groupbox.Anchor = 'Bottom'
	$ButtonDisconnect.Anchor = 'Bottom'
	$ButtonRunCustomCommand.Anchor = 'Bottom'
	
	#Make form sizable
	$FormO365AdministrationCenter.FormBorderStyle = 'Sizable'
}

	#Buttons

$ButtonDisconnect_Click = {
	
	$TextboxResults.Text = ""
	$textboxDetails.Text = ""
	
	#Disconnects O365 Session
	Get-PSSession | Remove-PSSession
	
	#Enables the connect to partner Button
	$PartnerConnectButton.Enabled = $True
	#Disabled the disconnect from partner button
	$ButtonDisconnect.Enabled = $False
	#Sets custom button text
	$PartnerConnectButton.Text = "Connect to Partner"
	#Sets the form name
	$FormO365AdministrationCenter.Text = "O365 Administration Center"
	#Enables the partner combobox
	$PartnerComboBox.Enabled = $True
	#Enables the connect to o365 button
	$ButtonConnectTo365.Enabled = $True
	#Clears the combobox
	#$PartnerComboBox.Items.clear()
		<# Creates a pop up box telling the user they are disconnected from the o365 session. This is commented out as it will show True every time as the command will never error out even if there 
		is no session to disconnect from #>
		#[void][System.Windows.Forms.MessageBox]::Show("You are disconnected from O365", "Message")
	}

<#
$buttonDisconnectFromPartner_Click = {
	$TextboxResults.Text = "Disconnecting from partner account..."
	Get-PSSession | Remove-PSSession
	Connect-MsolService -Credential $global:o365creds
	$exchangeSession = New-PSSession -Name MainAccount -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $global:o365creds -Authentication Basic -AllowRedirection
	Import-PSSession $exchangeSession
	#Disable Button
	$PartnerConnectButton.Enabled = $True
	$ButtonDisconnectFromPartner.Enabled = $False
	#Sets custom button text
	$PartnerConnectButton.Text = "Connect to Partner"
	$FormO365AdministrationCenter.Text = "O365 Administration Center"
	$PartnerComboBox.Enabled = $True
	$TextboxResults.Text = ""
	$textboxDetails.Text = ""

}
#>


$ButtonConnectTo365_Click = {
	try
	{
		$global:o365creds = (Get-Credential -Message "O365 Credentials")
		
		$TextboxResults.Text = "Creating implicit remoting module..."
		Connect-MsolService -Credential $global:o365creds
		
		$partnerTIDs = Get-MsolPartnerContract -All | Select-Object TenantID
		
		$domains = @()
		foreach ($TID in $partnerTIDs)
		{
			$domainName = Get-MsolDomain -TenantId $TID.TenantId | Where-Object { $_.IsDefault -eq 'True' }
			#$domainName = Get-MsolDomain -TenantId $TID.TenantId | Where-Object { $_.Name -notlike '*.onmicrosoft*' -and $_.Name -notlike '*.microsoftonline*' -and $_.Status -eq "Verified" }
			#$domainName = Get-MsolDomain -TenantId $TID.TenantId
			$domain = New-Object -TypeName PSObject
			$domain | Add-Member -Name 'Name' -MemberType NoteProperty -Value ($domainName.Name | Select-Object -First 1) #Deals with Tenants with multiple domain names asociated
			#$domain | Add-Member -Name 'Name' -MemberType NoteProperty -Value $domainName.Name
			$domain | Add-Member -Name TenantID -MemberType NoteProperty -Value $TID.TenantId
			$domains += $domain
		}
		
		#Loads Combobox with Tenants
		Load-ComboBox $PartnerComboBox $domains -DisplayMember Name
		
		#Connect to Exchange Online
		$exchangeSession = New-PSSession -Name MainAccount -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $global:o365creds -Authentication Basic -AllowRedirection
		Import-PSSession $exchangeSession -AllowClobber
		
		$TextboxResults.Text = ""
		
		$ButtonConnectTo365.Enabled = $False
		
		$ButtonConnectTo365.Text = "Connected to O365"
		
		$PartnerConnectButton.Enabled = $true
		
		$ButtonDisconnect.Enabled = $True
		
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show('You are not connected to O365, Please verify the correct username, password and that the PowerShell Execution policy is set to Unrestricted. Also make sure you have Microsoft Online Services Sign-In Assistant for IT Professionals installed. Check Help>Prerequisites for more info', "Error")
	}
}

$PartnerConnectButton_Click = {
	try
	{
		$URI = "https://outlook.office365.com/powershell-liveid?DelegatedOrg="+$PartnerComboBox.SelectedItem.Name
		Get-PSSession | Remove-PSSession
		#CONNECT TO EXCHANGE ONLINE
		$TextboxResults.Text = "Connecting to partner account..."
		$PartnerSession = New-PSSession -Name PartnerAccount -ConfigurationName Microsoft.Exchange -ConnectionUri $URI -Credential $global:o365creds -Authentication Basic -AllowRedirection
		Import-PSSession $PartnerSession -AllowClobber
		Connect-MsolService -Credential $global:o365creds
		#Disable Button
		$PartnerConnectButton.Enabled = $false
		#Sets custom button text
		$PartnerConnectButton.Text = "Connected to Partner"
		#Sets custom form text
		$FormO365AdministrationCenter.Text = "-Connected to"+ $PartnerComboBox.SelectedItem.Name+"-"
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		$PartnerComboBox.Enabled = $false
		#$ButtonDisconnectFromPartner.Enabled = $true
		$ButtonDisconnect.Enabled = $True
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("Unable to connect to partner", "Error")
	}
	
}

$ButtonExportToFile_Click={
	$SavedFile = Read-Host "Enter the Path for file (Eg. C:\DG.csv, C:\Users.txt, C:\output\info.doc)"
	try
	{
		$TextboxResults.Text | out-file $SavedFile
		[System.Windows.Forms.MessageBox]::Show("Saved $SavedFile", "Info")
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	} 
}

$ButtonRunCustomCommand_Click = {
	$userinput = $TextboxResults.text
		try
		{
			#Takes the user input to a variable and passes it to the shell
			$TextboxResults.text = Invoke-Expression $userinput | Out-String
		}
		Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
			[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
}



###USERS###

	#User General Items

$createOutOfOfficeAutoReplyForAUserToolStripMenuItem_Click = {
	$OOOautoreplyUser = Read-Host "What user is the Out Of Office auto reply for?"
	$OOOInternal = Read-Host "What is the Internal Message"
	$OOOExternal = Read-Host "What is the External Message"
	Try
	{
		$TextboxResults.Text = "Creating an out of office reply for $OOOautoreplyUser..."
		$textboxDetails.Text = "Set-MailboxAutoReplyConfiguration -Identity $OOOautoreplyUser -AutoReplyState Enabled -ExternalMessage $OOOExternal -InternalMessage $OOOInternal"
		Set-MailboxAutoReplyConfiguration -Identity $OOOautoreplyUser -AutoReplyState Enabled -ExternalMessage $OOOExternal -InternalMessage $OOOInternal
		$TextboxResults.Text = Get-MailboxAutoReplyConfiguration -Identity $OOOautoreplyUser | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getListOfUsersToolStripMenuItem_Click = {
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
		$TextboxResults.Text = "Getting list of users..."
		$textboxDetails.Text = "Get-MSOLUser | Sort-Object DisplayName |  Format-Table DisplayName, UserPrincipalName -AutoSize"
		$TextboxResults.text = Get-MSOLUser | Sort-Object DisplayName |  Format-Table DisplayName, UserPrincipalName -AutoSize | Out-String
		}
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue )
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting list of users..."
		$textboxDetails.Text = "Get-MSOLUser -TenantId $TenantText | Sort-Object DisplayName |  Format-Table DisplayName, UserPrincipalName -AutoSize "
		$TextboxResults.text = Get-MSOLUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Sort-Object DisplayName | Format-Table DisplayName, UserPrincipalName -AutoSize | Out-String
		}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("Could not get a list of users", "Error")
		}
}

$getDetailedInfoForAUserToolStripMenuItem_Click = {
	$DetailedInfoUser = Read-Host "Enter the UPN of the user you want more information about"
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting detailed info for $DetailedInfoUser..."
		$textboxDetails.Text = "Get-MsolUser -UserPrincipalName $DetailedInfoUser | Format-List"
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $DetailedInfoUser | Format-List | Out-String
	}
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting detailed info for $DetailedInfoUser..."
		$textboxDetails.Text = "Get-MsolUser -UserPrincipalName $DetailedInfoUser -TenantId $TenantText | Format-List"
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $DetailedInfoUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List | Out-String
		}
		Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
			[System.Windows.Forms.MessageBox]::Show("Could not get detailed info for $DetailedInfoUser", "Error")
		}
}

$changeUsersLoginNameToolStripMenuItem_Click = {
	$UserChangeUPN = Read-Host "What user would you like to change their login name for? Enter their UPN"
	$NewUserUPN = Read-Host "What would you like the new username to be?"
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
		$TextboxResults.Text = "Changing $UserChangeUPN UPN to $NewUserUPN..."
		$textboxDetails.Text = "Set-MsolUserPrincipalname -UserPrincipalName $UserChangeUPN -NewUserPrincipalName $NewUserUPN"
		Set-MsolUserPrincipalname -UserPrincipalName $UserChangeUPN -NewUserPrincipalName $NewUserUPN
		$TextboxResults.text = Get-MSOLUser -UserPrincipalName $NewUserUPN | Format-List UserPrincipalName | Out-String
		}
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Changing $UserChangeUPN UPN to $NewUserUPN..."
		$textboxDetails.Text = "Set-MsolUserPrincipalname -UserPrincipalName $UserChangeUPN -TenantId $TenantText -NewUserPrincipalName $NewUserUPN"
		Set-MsolUserPrincipalname -UserPrincipalName $UserChangeUPN -TenantId $PartnerComboBox.SelectedItem.TenantID -NewUserPrincipalName $NewUserUPN
		$TextboxResults.text = Get-MSOLUser -UserPrincipalName $NewUserUPN -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName | Out-String
		}
		Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
			[System.Windows.Forms.MessageBox]::Show("Could not change the login name for $UserChangeUPN", "Error")
		}
}

$deleteAUserToolStripMenuItem_Click = {
	$DeleteUser = Read-Host "Enter the UPN of the user you want to delete"
		#What to do if connected to main o365 account
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
		$TextboxResults.Text = "Deleting $DeleteUser..."
		$textboxDetails.Text = "Remove-MsolUser –UserPrincipalName $DeleteUser"
		$TextboxResults.text = Remove-MsolUser –UserPrincipalName $DeleteUser | Format-List UserPrincipalName | Out-String
		}
		#What to do if connected to partner account
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Deleting $DeleteUser..."
		$textboxDetails.Text = "Remove-MsolUser –UserPrincipalName $DeleteUser -TenantId $TenantText"
		$TextboxResults.text = Remove-MsolUser –UserPrincipalName $DeleteUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName | Out-String
		}
		Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
			[System.Windows.Forms.MessageBox]::Show("Could not delete $DeleteUser", "Error")
		}	
}

$createANewUserToolStripMenuItem_Click = {
	$Firstname = Read-Host "Enter the First Name for the new user"
	$LastName = Read-Host "Enter the Last Name for the new user"
	$DisplayName = Read-Host "Enter the Display Name for the new user"
	$NewUser = Read-Host "Enter the UPN for the new user"
		#What to do if connected to main o365 account
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
		$TextboxResults.Text = "Creating user $NewUser..."
		$textboxDetails.Text = "New-MsolUser -UserPrincipalName $NewUser -FirstName $Firstname -LastName $LastName -DisplayName $DisplayName"
		$TextboxResults.text = New-MsolUser -UserPrincipalName $NewUser -FirstName $Firstname -LastName $LastName -DisplayName $DisplayName | Format-List | Out-String
		}
		#What to do if connected to partner account
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Creating user $NewUser..."
		$textboxDetails.Text = "New-MsolUser -TenantId $TenantText -UserPrincipalName $NewUser -FirstName $Firstname -LastName $LastName -DisplayName $DisplayName"
		$TextboxResults.text = New-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $NewUser -FirstName $Firstname -LastName $LastName -DisplayName $DisplayName | Format-List | Out-String
		}
		Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
			[System.Windows.Forms.MessageBox]::Show("Could not create the new user $newuser", "Error")
		}
}

$disableUserAccountToolStripMenuItem_Click = {
	$BlockUser = Read-Host "Enter the UPN of the user you want to disable"
		#What to do if connected to main o365 account
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
		$TextboxResults.Text = "Disabling $BlockUser..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $BlockUser -blockcredential `$True"
		Set-MsolUser -UserPrincipalName $BlockUser -blockcredential $True 
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $BlockUser | Format-List DisplayName, BlockCredential | Out-String
		}
		#What to do if connected to partner account
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Disabling $BlockUser..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $BlockUser -blockcredential `$True -TenantId $TenantText"
		Set-MsolUser -UserPrincipalName $BlockUser -blockcredential $True -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $BlockUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List DisplayName, BlockCredential | Out-String
		}
		Else
		{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
			[System.Windows.Forms.MessageBox]::Show("Could not disable $BlockUser", "Error")
		}
}

$enableAccountToolStripMenuItem_Click = {
	$EnableUser = Read-Host "Enter the UPN of the user you want to enable"
		#What to do if connected to main o365 account
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
		$TextboxResults.Text = "Enabling $EnableUser..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $EnableUser -blockcredential `$False"
		Set-MsolUser -UserPrincipalName $EnableUser -blockcredential $False
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $EnableUser | Format-List DisplayName, BlockCredential | Out-String
		}
		#What to do if connected to partner account
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Enabling $EnableUser..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $EnableUser -blockcredential `$False -TenantId $TenantText"
		Set-MsolUser -UserPrincipalName $EnableUser -blockcredential $False -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $EnableUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List DisplayName, BlockCredential | Out-String
		}
		Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("Could enable $EnableUser", "Error")
		}
}

	#Quota

$getUserQuotaToolStripMenuItem_Click={
	$QuotaUser = Read-Host "Enter the Email of the user you want to view Quota information for"
	try
	{
		$TextboxResults.Text = "Getting user quota for $QuotaUser..."
		$textboxDetails.Text = "Get-Mailbox $QuotaUser | Format-List DisplayName, UserPrincipalName, *Quota"
		$TextboxResults.text = Get-Mailbox $QuotaUser | Format-List DisplayName, UserPrincipalName, *Quota | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getAllUsersQuotaToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting quota for all users..."
		$textboxDetails.Text = "Get-Mailbox | Format-List DisplayName, UserPrincipalName, *Quota -AutoSize"
		$TextboxResults.text = Get-Mailbox | Sort-Object DisplayName | Format-Table DisplayName, UserPrincipalName, *Quota -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$setUserMailboxQuotaToolStripMenuItem_Click = {
	$MailboxSetQuota = Read-Host "Enter the UPN of the user you want to edit quota for"
	$ProhibitSendReceiveQuota = Read-Host "Enter (GB) the ProhibitSendReceiveQuota value (EX: 50GB) Max:50GB"
	$ProhibitSendQuota = Read-Host "Enter (GB) the ProhibitSendQuota value (EX: 48GB) Max:50GB"
	$IssueWarningQuota = Read-Host "Enter (GB) theIssueWarningQuota value (EX: 45GB) Max:50GB"
	Try
	{
		$TextboxResults.Text = "Setting quota for $MailboxSetQuota... "
		$textboxDetails.Text = "Set-Mailbox $MailboxSetQuota -ProhibitSendReceiveQuota $ProhibitSendReceiveQuota -ProhibitSendQuota $ProhibitSendQuota -IssueWarningQuota $IssueWarningQuota"
		Set-Mailbox $MailboxSetQuota -ProhibitSendReceiveQuota $ProhibitSendReceiveQuota -ProhibitSendQuota $ProhibitSendQuota -IssueWarningQuota $IssueWarningQuota
		$TextboxResults.text = Get-Mailbox $MailboxSetQuota | Format-List DisplayName, UserPrincipalName, *Quota | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$setMailboxQuotaForAllToolStripMenuItem_Click = {
	$ProhibitSendReceiveQuota2 = Read-Host "Enter (GB) the ProhibitSendReceiveQuota value (EX: 50GB) Max:50GB"
	$ProhibitSendQuota2 = Read-Host "Enter (GB) the ProhibitSendQuota value (EX: 48GB) Max:50GB"
	$IssueWarningQuota2 = Read-Host "Enter (GB) theIssueWarningQuota value (EX: 45GB) Max:50GB"
	Try
	{
		$TextboxResults.Text = "Setting quota for all... "
		$textboxDetails.Text = "Get-Mailbox | Set-Mailbox -ProhibitSendReceiveQuota $ProhibitSendReceiveQuota2 -ProhibitSendQuota $ProhibitSendQuota2 -IssueWarningQuota $IssueWarningQuota2"
		Get-Mailbox | Set-Mailbox -ProhibitSendReceiveQuota $ProhibitSendReceiveQuota2 -ProhibitSendQuota $ProhibitSendQuota2 -IssueWarningQuota $IssueWarningQuota2
		$TextboxResults.text = Get-Mailbox | Format-List DisplayName, UserPrincipalName, *Quota | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}


	#Licenses

$getLicensedUsersToolStripMenuItem_Click={
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all users with a license..."
		$textboxDetails.Text = "Get-MsolUser | Where-Object { `$_.isLicensed -eq `$TRUE } | Sort-Object DisplayName | Format-Table DisplayName, Licenses -AutoSize"
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.isLicensed -eq $True } | Sort-Object DisplayName | Format-Table DisplayName, Licenses -AutoSize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting all users with a license..."
		$textboxDetails.Text = "Get-MsolUser -TenantID $TenantText | Where-Object { `$_.isLicensed -eq `$TRUE } | Sort-Object DisplayName | Format-Table DisplayName, Licenses -AutoSize"
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.isLicensed -eq $True } | Sort-Object DisplayName | Format-Table DisplayName, Licenses -AutoSize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("Could not get license information", "Error")
	}
}

$displayAllUsersWithoutALicenseToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all users without a license..."
		$textboxDetails.Text = "Get-MsolUser | Where-Object { `$_.isLicensed -like `$False } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName -AutoSize"
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.isLicensed -like "False" } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName -AutoSize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting all users without a license..."
		$textboxDetails.Text = "Get-MsolUser -TenantId $TenantText | Where-Object { `$_.isLicensed -like `$False } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName -AutoSize"
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.isLicensed -like $False } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName -AutoSize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("Could not display users without a license", "Error")
	}
	
}

$removeAllUnlicensedUsersToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Removing all users without a license..."
		$textboxDetails.Text = "Get-MsolUser | Where-Object { `$_.isLicensed -ne `$True } | Remove-MsolUser -Force"
		Get-MsolUser | Where-Object { $_.isLicensed -ne $True } | Remove-MsolUser -Force
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.isLicensed -like $True } | Format-List UserPrincipalName | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Removing all users without a license..."
		$textboxDetails.Text = "Get-MsolUser -TenantId $TenantText | Where-Object { `$_.isLicensed -ne `$True } | Remove-MsolUser -Force"
		Get-MsolUser -all -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.isLicensed -ne $True } | Remove-MsolUser -Force -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.isLicensed -like $True } | Format-List UserPrincipalName | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("Could not remove all unlicensed users", "Error")
	}
}

$displayAllLicenseInfoToolStripMenuItem_Click = {
#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all license information..."
		$textboxDetails.Text = "Get-MsolAccountSku | Format-Table"
		$TextboxResults.text = Get-MsolAccountSku | Select-Object -Property AccountSkuId, ActiveUnits, WarningUnits, ConsumedUnits, @{
			Name = 'Unused'
			Expression = {
				$_.ActiveUnits - $_.ConsumedUnits
			}
		} | Sort-Object Unused | Format-Table AccountSkuId, ActiveUnits, WarningUnits, ConsumedUnits, Unused -AutoSize | Out-String 
	}
    #What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting all license information..."
		$textboxDetails.Text = "Get-MsolAccountSku -TenantId $TenantText | Format-Table"
		$TextboxResults.text = Get-MsolAccountSku -TenantId $PartnerComboBox.SelectedItem.TenantID | Select-Object -Property AccountSkuId, ActiveUnits, WarningUnits, ConsumedUnits, @{
			Name = 'Unused'
			Expression = {
				$_.ActiveUnits - $_.ConsumedUnits
			}
		} | Sort-Object Unused | Format-Table AccountSkuId, ActiveUnits, WarningUnits, ConsumedUnits, Unused -AutoSize | Out-String 
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$addALicenseToAUserToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$LicenseUserAdd = Read-Host "Enter the User Principal Name of the User you want to license"
		$LicenseUserAddLocation = Read-Host "Enter the 2 digit location code for the user. Example: US"
		$TextboxResults.text = Get-MsolAccountSku | Format-Table | Out-String
		$LicenseType = Read-Host "Enter the AccountSku of the License you want to assign to this user"
		$TextboxResults.Text = "Adding $LicenseType license to $LicenseUserAdd..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $LicenseUserAdd –UsageLocation $LicenseUserAddLocation
		Set-MsolUserLicense -UserPrincipalName $LicenseUserAdd -AddLicenses $LicenseType"
		Set-MsolUser -UserPrincipalName $LicenseUserAdd –UsageLocation $LicenseUserAddLocation
		Set-MsolUserLicense -UserPrincipalName $LicenseUserAdd -AddLicenses $LicenseType
		$TextboxResults.Text = Get-MsolUser -UserPrincipalName $LicenseUserAdd | Format-List DisplayName, Licenses | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$LicenseUserAdd = Read-Host "Enter the User Principal Name of the User you want to license"
		$LicenseUserAddLocation = Read-Host "Enter the 2 digit location code for the user. Example: US"
		$TextboxResults.text = Get-MsolAccountSku -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-Table | Out-String
		$LicenseType = Read-Host "Enter the AccountSku of the License you want to assign to this user"
		$TextboxResults.Text = "Adding $LicenseType license to $LicenseUserAdd..."
		$textboxDetails.Text = "Set-MsolUser -TenantId $TenantText -UserPrincipalName $LicenseUserAdd –UsageLocation $LicenseUserAddLocation
		Set-MsolUserLicense -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $LicenseUserAdd -AddLicenses $LicenseType"
		Set-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $LicenseUserAdd –UsageLocation $LicenseUserAddLocation
		Set-MsolUserLicense -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $LicenseUserAdd -AddLicenses $LicenseType
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $LicenseUserAdd | Format-List DisplayName, Licenses | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$removeLicenseFromAUserToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$RemoveLicenseFromUser = Read-Host "Enter the User Principal Name of the user you want to remove a license from"
		$TextboxResults.Text = Get-MsolUser -UserPrincipalName $RemoveLicenseFromUser | Format-List DisplayName, Licenses | Out-String
		$RemoveLicenseType = Read-Host "Enter the AccountSku of the license you want to remove"
		$TextboxResults.Text = "Removing the $RemoveLicenseType license from $RemoveLicenseFromUser..."
		$textboxDetails.Text = "Set-MsolUserLicense -UserPrincipalName $RemoveLicenseFromUser -RemoveLicenses $RemoveLicenseType"
		Set-MsolUserLicense -UserPrincipalName $RemoveLicenseFromUser -RemoveLicenses $RemoveLicenseType
		$TextboxResults.Text = Get-MsolUser -UserPrincipalName $RemoveLicenseFromUser | Format-List DisplayName, Licenses | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$RemoveLicenseFromUser = Read-Host "Enter the User Principal Name of the user you want to remove a license from"
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $RemoveLicenseFromUser | Format-List DisplayName, Licenses | Out-String
		$RemoveLicenseType = Read-Host "Enter the AccountSku of the license you want to remove"
		$TextboxResults.Text = "Removing the $RemoveLicenseType license from $RemoveLicenseFromUser..."
		$textboxDetails.Text = "Set-MsolUserLicense -TenantId $TenantText -UserPrincipalName $RemoveLicenseFromUser -RemoveLicenses $RemoveLicenseType"
		Set-MsolUserLicense -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $RemoveLicenseFromUser -RemoveLicenses $RemoveLicenseType
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $RemoveLicenseFromUser | Format-List DisplayName, Licenses | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Calendar Items

$addCalendarPermissionsToolStripMenuItem_Click = {
	$Calendaruser = Read-Host "Enter the UPN of the user whose Calendar you want to give access to"
	$Calendaruser2 = Read-Host "Enter the UPN of the user who you want to give access to"
	$TextboxResults.text = "Calendar Permissions: 
Owner
PublishingEditor
Editor
PublishingAuthor
Author
NonEditingAuthor
Reviewer
Contributor
AvailabilityOnly
LimitedDetails"
	$level = Read-Host "Access Level?"
	try
	{
		$TextboxResults.Text = "Adding $Calendaruser2 to $Calendaruser calender with $level permissions..."
		$textboxDetails.Text = "Add-MailboxFolderPermission -Identity ${Calendaruser}:\calendar -user $Calendaruser2 -AccessRights $level"
		Remove-MailboxFolderPermission -identity ${Calendaruser}:\calendar -user $Calendaruser2 -Confirm:$False
		Add-MailboxFolderPermission -Identity ${Calendaruser}:\calendar -user $Calendaruser2 -AccessRights $level
		$TextboxResults.Text = Get-MailboxFolderPermission -Identity ${Calendaruser}:\calendar | Sort-Object User, AccessRights | Format-Table User, AccessRights, Identity, FolderName, IsValid -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$viewUsersCalendarPermissionsToolStripMenuItem_Click = {
	$CalUserPermissions = Read-Host "What user would you like calendar permissions for?"
	Try
	{
		$TextboxResults.Text = "Getting $CalUserPermissions calendar permissions..."
		$textboxDetails.Text = "Get-MailboxFolderPermission -Identity ${CalUserPermissions}:\calendar | Sort-Object User, AccessRights | Format-Table User, AccessRights, Identity, FolderName, IsValid -AutoSize"
		$TextboxResults.Text = Get-MailboxFolderPermission -Identity ${CalUserPermissions}:\calendar | Sort-Object User, AccessRights | Format-Table User, AccessRights, Identity, FolderName, IsValid -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$addsingsUsersPermissionToAllCalrndarToolStripMenuItem_Click = {
	$MasterUser = Read-Host "Enter the UPN of the user you want permission on all users calendars"
	$TextboxResults.text = "Calendar Permissions: 
Owner
PublishingEditor
Editor
PublishingAuthor
Author
NonEditingAuthor
Reviewer
Contributor
AvailabilityOnly
LimitedDetails"
$level2 = Read-Host "Access Level?"
	try
	{
		$TextboxResults.Text = "Adding $MasterUser to everyones calendars with $level2 permissions..."
		$textboxDetails.Text = "Get-Mailbox | Select-Object -ExpandProperty Alias
Foreach (`$user in `$users) { Add-MailboxFolderPermission `${user}:\Calendar -user $MasterUser -accessrights $level2 }﻿"
		$users = Get-Mailbox | Select-Object -ExpandProperty Alias
		Foreach ($user in $users) { Add-MailboxFolderPermission ${user}:\Calendar -user $MasterUser -accessrights $level2 }﻿
	}
	catch
	{
		#[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	$TextboxResults.Text = ""

}

$removeAUserFromAllCalendarsToolStripMenuItem_Click = {
	$RemoveUserFromAll = Read-Host "Enter the UPN of the user you want to remove from all calendars"
	try
	{
		$TextboxResults.Text = "Removing $RemoveUserFromAll from all users calendar..."
		$textboxDetails.Text = "`$users = Get-Mailbox | Select-Object -ExpandProperty Alias
Foreach (`$user in `$users) { Remove-MailboxFolderPermission `${user}:\Calendar -user $RemoveUserFromAll -Confirm:`$false}﻿"
		$users = Get-Mailbox | Select-Object -ExpandProperty Alias
		Foreach ($user in $users) { Remove-MailboxFolderPermission ${user}:\Calendar -user $RemoveUserFromAll -Confirm:$false }﻿
	}
	catch
	{
		#[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	$TextboxResults.Text = ""
}

$removeAUserFromSomesonsCalendarToolStripMenuItem_Click = {
	$Calendaruserremove = Read-Host "Enter the UPN of the user whose calendar you want to remove access to"
	$Calendaruser2remove = Read-Host "Enter the UPN of the user who you want to remove access"
	try
	{
		$TextboxResults.Text = "Removing $Calendaruser2remove from $Calendaruserremove calendar..."
		$textboxDetails.Text = "Remove-MailboxFolderPermission -Identity ${Calendaruserremove}:\calendar -user $Calendaruser2remove"
		Remove-MailboxFolderPermission -Identity ${Calendaruserremove}:\calendar -user $Calendaruser2remove
		$TextboxResults.Text = Get-MailboxFolderPermission -Identity ${Calendaruserremove}:\calendar | Sort-Object User, AccessRights | Format-Table User, AccessRights, Identity, FolderName, IsValid -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Clutter

$disableClutterForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Disabling Clutter for all users..."
		$textboxDetails.Text = "Get-Mailbox | Set-Clutter -Enable `$false | Format-List MailboxIdentity, IsEnabled"
		$TextboxResults.text = Get-Mailbox | Set-Clutter -Enable $false | Format-List MailboxIdentity, IsEnabled | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$enableClutterForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Enabling Clutter for all users..."
		$textboxDetails.Text = "Get-Mailbox | Set-Clutter -Enable `$True | Format-List MailboxIdentity, IsEnabled"
		$TextboxResults.text = Get-Mailbox | Set-Clutter -Enable $True | Format-List MailboxIdentity, IsEnabled | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$enableClutterForAUserToolStripMenuItem_Click = {
	$UserEnableClutter = Read-Host "Which user would you like to enable Clutter for?"
	try
	{
		$TextboxResults.Text = "Enabling Clutter for $UserEnableClutter..."
		$textboxDetails.Text = "Get-Mailbox $UserEnableClutter | Set-Clutter -Enable `$True"
		$TextboxResults.text = Get-Mailbox $UserEnableClutter | Set-Clutter -Enable $True | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$disableClutterForAUserToolStripMenuItem_Click = {
	$UserDisableClutter = Read-Host "Which user would you like to disable Clutter for?"
	try
	{
		$TextboxResults.Text = "Disabling Clutter for $UserDisableClutter..."
		$textboxDetails.Text = "Get-Mailbox $UserDisableClutter | Set-Clutter -Enable `$False"
		$TextboxResults.text = Get-Mailbox $UserDisableClutter | Set-Clutter -Enable $False | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getClutterInfoForAUserToolStripMenuItem_Click = {
	$GetCluterInfoUser = Read-Host "What user would you like to view Clutter information about?"
	Try
	{
		$TextboxResults.Text = "Getting Clutter information for $GetCluterInfoUser..."
		$textboxDetails.Text = "Get-Clutter -Identity $GetCluterInfoUser"
		$TextboxResults.Text = Get-Clutter -Identity $GetCluterInfoUser | Format-List MailboxIdentity, IsEnabled | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Recycle Bin

$displayAllDeletedUsersToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all deleted users..."
		$textboxDetails.Text = "Get-MsolUser -ReturnDeletedUsers |  Sort-Object UserprincipalName | Format-Table UserPrincipalName, ObjectID -Autosize "
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Sort-Object UserprincipalName | Format-Table UserPrincipalName, ObjectID -Autosize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting all deleted users..."
		$textboxDetails.Text = "Get-MsolUser -TenantId $TenantText -ReturnDeletedUsers |  Sort-Object UserprincipalName | Format-Table UserPrincipalName, ObjectID -Autosize "
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -ReturnDeletedUsers | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, ObjectID -AutoSize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("Could not display users without a license", "Error")
	}
}

$deleteAllUsersInRecycleBinToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Deleting all users in the recycle bin..."
		$textboxDetails.Text = "Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin –Force"
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin –Force | Format-List | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Deleting all users in the recycle bin..."
		$textboxDetails.Text = "Get-MsolUser -TenantId $TenantText -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin –Force"
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin –Force | Format-List | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("Could not display users without a license", "Error")
	}
	
}

$deleteSpecificUsersInRecycleBinToolStripMenuItem_Click = {
	$DeletedUserRecycleBin = Read-Host "Please enter the User Principal Name of the user you want to permanently delete"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Deleting  $DeletedUserRecycleBin from the recycle bin..."
		$textboxDetails.Text = "Remove-MsolUser -UserPrincipalName $DeletedUserRecycleBin -RemoveFromRecycleBin -Force"
		Remove-MsolUser -UserPrincipalName $DeletedUserRecycleBin -RemoveFromRecycleBin -Force
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Sort-Object UserprincipalName | Format-Table UserPrincipalName, ObjectID -Autosize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Deleting  $DeletedUserRecycleBin from the recycle bin..."
		$textboxDetails.Text = "Remove-MsolUser -TenantId $TenantText -UserPrincipalName $DeletedUserRecycleBin -RemoveFromRecycleBin -Force"
		Remove-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $DeletedUserRecycleBin -RemoveFromRecycleBin -Force
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -ReturnDeletedUsers | Sort-Object UserprincipalName | Format-Table UserPrincipalName, ObjectID -Autosize  | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("Could not display users without a license", "Error")
	}
}

$restoreDeletedUserToolStripMenuItem_Click = {
	$RestoredUserFromRecycleBin = Read-Host "Enter the User Principal Name of the user you want to restore"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Restoring $RestoredUserFromRecycleBin from the recycle bin..."
		$textboxDetails.Text = "Restore-MsolUser –UserPrincipalName $RestoredUserFromRecycleBin -AutoReconcileProxyConflicts"
		Restore-MsolUser –UserPrincipalName $RestoredUserFromRecycleBin -AutoReconcileProxyConflicts
		$TextboxResults.Text = "Getting list of deleted users"
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Sort-Object UserprincipalName | Format-Table UserPrincipalName, ObjectID -Autosize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Restoring $RestoredUserFromRecycleBin from the recycle bin..."
		$textboxDetails.Text = "Restore-MsolUser -TenantId $TenantText –UserPrincipalName $RestoredUserFromRecycleBin -AutoReconcileProxyConflicts"
		Restore-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID –UserPrincipalName $RestoredUserFromRecycleBin -AutoReconcileProxyConflicts
		$TextboxResults.Text = "Getting list of deleted users"
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -ReturnDeletedUsers | Sort-Object UserprincipalName | Format-Table UserPrincipalName, ObjectID -Autosize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$restoreAllDeletedUsersToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Restoring all deleted users..."
		$textboxDetails.Text = "Get-MsolUser -ReturnDeletedUsers | Restore-MsolUser"
		Get-MsolUser -ReturnDeletedUsers | Restore-MsolUser
		$TextboxResults.Text = "Users that were deleted have now been restored"
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Restoring all deleted users..."
		$textboxDetails.Text = "Get-MsolUser -ReturnDeletedUsers -TenantID $TenantText | Restore-MsolUser"
		Get-MsolUser -ReturnDeletedUsers -TenantId $PartnerComboBox.SelectedItem.TenantID | Restore-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.Text = "Users that were deleted have now been restored"
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Quarentine

$getQuarantineBetweenDatesToolStripMenuItem_Click = {
	$StartDateQuarentine = Read-Host "Enter the beginning date. (Format MM/DD/YYYY)"
	$EndDateQuarentine = Read-Host "Enter the end date. (Format MM/DD/YYYY)"
	try
	{
		$TextboxResults.Text = "Getting quarantine between $StartDateQuarentine and $EndDateQuarentine..."
		$textboxDetails.Text = "Get-QuarantineMessage -StartReceivedDate $StartDateQuarentine -EndReceivedDate $EndDateQuarentine | Format-List ReceivedTime, SenderAddress, RecipientAddress, Subject, Size, Type, Expires, QuarantinedUser, ReleasedUser, Direction "
		$TextboxResults.Text = Get-QuarantineMessage -StartReceivedDate $StartDateQuarentine -EndReceivedDate $EndDateQuarentine | Format-List ReceivedTime, SenderAddress, RecipientAddress, Subject, Size, Type, Expires, QuarantinedUser, ReleasedUser, Direction | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getQuarantineFromASpecificUserToolStripMenuItem_Click = {
	$QuarentineFromUser = Read-Host "Enter the email address you want to see quarentine from"
	try
	{
		$TextboxResults.Text = "Getting quarantine sent from $QuarentineFromUser ..."
		$textboxDetails.Text = "Get-QuarantineMessage -SenderAddress $QuarentineFromUser | Format-List ReceivedTime, SenderAddress, RecipientAddress, Subject, Size, Type, Expires, QuarantinedUser, ReleasedUser, Direction"
		$TextboxResults.Text = Get-QuarantineMessage -SenderAddress $QuarentineFromUser | Format-List ReceivedTime, SenderAddress, RecipientAddress, Subject, Size, Type, Expires, QuarantinedUser, ReleasedUser, Direction | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getQuarantineToASpecificUserToolStripMenuItem_Click = {
	$QuarentineInfoForUser = Read-Host "Enter the email of the user you want to view quarantine for"
	try
	{
		$TextboxResults.Text = "Getting quarantine sent to $QuarentineInfoForUser..."
		$textboxDetails.Text = " Get-QuarantineMessage -RecipientAddress $QuarentineInfoForUser | Format-List ReceivedTime, SenderAddress, Subject, Size, Type, Expires, QuarantinedUser, ReleasedUser, Direction"
		$TextboxResults.Text = Get-QuarantineMessage -RecipientAddress $QuarentineInfoForUser | Format-List ReceivedTime, SenderAddress, Subject, Size, Type, Expires, QuarantinedUser, ReleasedUser, Direction | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Passwords

$enableStrongPasswordForAUserToolStripMenuItem_Click = {
	$UserEnableStrongPasswords = Read-Host "Enter the User Principal Name of the user you want to enable strong password policy for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Enabling strong password policy for $UserEnableStrongPasswords..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $UserEnableStrongPasswords -StrongPasswordRequired `$True"
		Set-MsolUser -UserPrincipalName $UserEnableStrongPasswords -StrongPasswordRequired $True
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $UserEnableStrongPasswords | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
		
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Enabling strong password policy for $UserEnableStrongPasswords..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $UserEnableStrongPasswords -StrongPasswordRequired `$True -TenantId $TenantText"
		Set-MsolUser -UserPrincipalName $UserEnableStrongPasswords -StrongPasswordRequired $True -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $UserEnableStrongPasswords -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
		
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getAllUsersStrongPasswordPolicyInfoToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting strong password policy for all users..."
		$textboxDetails.Text = "Get-MsolUser | Sort-Object DisplayName | Format-Table DisplayName, strongpasswordrequired -AutoSize"
		$TextboxResults.text = Get-MsolUser | Sort-Object DisplayName | Format-Table  DisplayName, strongpasswordrequired -AutoSize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting strong password policy for all users..."
		$textboxDetails.Text = "Get-MsolUser -TenantId $TenantText | Sort-Object DisplayName | Format-Table  DisplayName, strongpasswordrequired -AutoSize"
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Sort-Object DisplayName | Format-Table DisplayName, strongpasswordrequired -AutoSize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$disableStrongPasswordsForAUserToolStripMenuItem_Click = {
	$UserdisableStrongPasswords = Read-Host "Enter the User Principal Name of the user you want to disable strong password policy for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Disabling strong password policy for $UserdisableStrongPasswords..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $UserdisableStrongPasswords -StrongPasswordRequired `$False"
		Set-MsolUser -UserPrincipalName $UserdisableStrongPasswords -StrongPasswordRequired $False
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $UserdisableStrongPasswords | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
		
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Disabling strong password policy for $UserdisableStrongPasswords..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $UserdisableStrongPasswords -StrongPasswordRequired `$False -TenantID $TenantText"
		Set-MsolUser -UserPrincipalName $UserdisableStrongPasswords -StrongPasswordRequired $False -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $UserdisableStrongPasswords -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$enableStrongPasswordsForAllToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Enabling strong password policy for all users..."
		$textboxDetails.Text = "Get-MsolUser | Set-MsolUser -StrongPasswordRequired `$True"
		Get-MsolUser | Set-MsolUser -StrongPasswordRequired $True
		$TextboxResults.text = Get-MsolUser | Sort-Object DisplayName | Format-Table  DisplayName, strongpasswordrequired -AutoSize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Enabling strong password policy for all users..."
		$textboxDetails.Text = "Get-MsolUser | Set-MsolUser -StrongPasswordRequired -TenantId $TenantText `$True"
		Get-MsolUser | Set-MsolUser -StrongPasswordRequired $True -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Sort-Object DisplayName | Format-Table  DisplayName, strongpasswordrequired -AutoSize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$disableStrongPasswordsForAllToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Disabling strong password policy for all users..."
		$textboxDetails.Text = "Get-MsolUser | Set-MsolUser -StrongPasswordRequired `$False"
		Get-MsolUser | Set-MsolUser -StrongPasswordRequired $False
		$TextboxResults.text = Get-MsolUser | Sort-Object DisplayName | Format-Table  DisplayName, strongpasswordrequired -AutoSize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Disabling strong password policy for all users..."
		$textboxDetails.Text = "Get-MsolUser | Set-MsolUser -StrongPasswordRequired `$False -TenantId $TenantText"
		Get-MsolUser | Set-MsolUser -StrongPasswordRequired $False -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Sort-Object DisplayName | Format-Table  DisplayName, strongpasswordrequired -AutoSize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$resetPasswordForAUserToolStripMenuItem1_Click = {
	$ResetPasswordUser = Read-Host "Who user would you like to reset the password for?"
	$NewPassword = Read-Host "What would you like the new password to be?"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Resetting $ResetPasswordUser password to $NewPassword..."
		$textboxDetails.Text = "Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser –NewPassword $NewPassword -ForceChangePassword `$False"
		Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser –NewPassword $NewPassword -ForceChangePassword $False
		$TextboxResults.Text = "The password for $ResetPasswordUser has been set to $NewPassword"
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Resetting $ResetPasswordUser password to $NewPassword..."
		$textboxDetails.Text = "Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser –NewPassword $NewPassword -ForceChangePassword `$False -TenantId $TenantText"
		Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser –NewPassword $NewPassword -ForceChangePassword $False -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.Text = "The password for $ResetPasswordUser has been set to $NewPassword"
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$setPasswordToNeverExpireForAllToolStripMenuItem1_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to never expire for all..."
		$textboxDetails.Text = "Get-MsolUser | Set-MsolUser –PasswordNeverExpires `$True"
		Get-MsolUser | Set-MsolUser –PasswordNeverExpires $True
		$TextboxResults.text = Get-MSOLUser | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires -AutoSize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Setting password to never expire for all..."
		$textboxDetails.Text = "Get-MsolUser | Set-MsolUser –PasswordNeverExpires `$True -TenantId $TenantText"
		Get-MsolUser | Set-MsolUser –PasswordNeverExpires $True -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MSOLUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires -AutoSize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$setPasswordToExpireForAllToolStripMenuItem1_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to expire for all..."
		$textboxDetails.Text = "Get-MsolUser | Set-MsolUser –PasswordNeverExpires `$False"
		Get-MsolUser | Set-MsolUser –PasswordNeverExpires $False
		$TextboxResults.text = Get-MSOLUser | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires -AutoSize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Setting password to expire for all..."
		$textboxDetails.Text = "Get-MsolUser | Set-MsolUser –PasswordNeverExpires `$False -TenantId $TenantText"
		Get-MsolUser | Set-MsolUser –PasswordNeverExpires $False -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MSOLUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires -AutoSize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$resetPasswordForAllToolStripMenuItem_Click = {
	$SetPasswordforAll = Read-Host "What password would you like to set for all users?"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Resetting all users passwords to $SetPasswordforAll..."
		$textboxDetails.Text = "Get-MsolUser | ForEach-Object{ 
Set-MsolUserPassword -userPrincipalName `$_.UserPrincipalName –NewPassword $SetPasswordforAll -ForceChangePassword `$False }"
		Get-MsolUser | ForEach-Object{ Set-MsolUserPassword -userPrincipalName $_.UserPrincipalName –NewPassword $SetPasswordforAll -ForceChangePassword $False }
		$TextboxResults.Text = "Password for all users has been set to $SetPasswordforAll"
		
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Resetting all users passwords to $SetPasswordforAll..."
		$textboxDetails.Text = "Get-MsolUser | ForEach-Object{ 
Set-MsolUserPassword -TenantId $TenantText -userPrincipalName `$_.UserPrincipalName –NewPassword $SetPasswordforAll -ForceChangePassword `$False }"
		Get-MsolUser | ForEach-Object{ Set-MsolUserPassword -TenantId $PartnerComboBox.SelectedItem.TenantID -userPrincipalName $_.UserPrincipalName –NewPassword $SetPasswordforAll -ForceChangePassword $False }
		$TextboxResults.Text = "Password for all users has been set to $SetPasswordforAll"
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$setATemporaryPasswordForAllToolStripMenuItem_Click = {
	$SetTempPasswordforAll = Read-Host "What password would you like to set for all users?"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting $SetTempPasswordforAll as the temporary password for all users..."
		$textboxDetails.Text = "Get-MsolUser | Set-MsolUserPassword –NewPassword $SetTempPasswordforAll -ForceChangePassword `$True"
		Get-MsolUser | Set-MsolUserPassword –NewPassword $SetTempPasswordforAll -ForceChangePassword $True
		$TextboxResults.Text = "Temporary password has been set to $SetTempPasswordforAll Please note that users will be prompted to change it upon logon"
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Setting $SetTempPasswordforAll as the temporary password for all users..."
		$textboxDetails.Text = "Get-MsolUser | Set-MsolUserPassword -TenantId $TenantText –NewPassword $SetTempPasswordforAll -ForceChangePassword `$True"
		Get-MsolUser | Set-MsolUserPassword -TenantId $PartnerComboBox.SelectedItem.TenantID –NewPassword $SetTempPasswordforAll -ForceChangePassword $True
		$TextboxResults.Text = "Temporary password has been set to $SetTempPasswordforAll Please note that users will be prompted to change it upon logon"
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$TemporaryPasswordForAUserToolStripMenuItem_Click = {
	$ResetPasswordUser2 = Read-Host "Who user would you like to reset the password for?"
	$NewPassword2 = Read-Host "What would you like the new password to be?"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting $NewPassword2 as the temporary password for $ResetPasswordUser2..."
		$textboxDetails.Text = "Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser2 –NewPassword $NewPassword2 -ForceChangePassword `$True"
		Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser2 –NewPassword $NewPassword2 -ForceChangePassword $True
		$TextboxResults.Text = "Temporary password has been set to $NewPassword2 Please note that $ResetPasswordUser2 will be prompted to change it upon logon"
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Setting $NewPassword2 as the temporary password for $ResetPasswordUser2..."
		$textboxDetails.Text = "Set-MsolUserPassword -TenantId $TenantText –UserPrincipalName $ResetPasswordUser2 –NewPassword $NewPassword2 -ForceChangePassword `$True"
		Set-MsolUserPassword -TenantId $PartnerComboBox.SelectedItem.TenantID –UserPrincipalName $ResetPasswordUser2 –NewPassword $NewPassword2 -ForceChangePassword $True
		$TextboxResults.Text = "Temporary password has been set to $NewPassword2 Please note that $ResetPasswordUser2 will be prompted to change it upon logon"
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getPasswordResetDateForAUserToolStripMenuItem_Click = {
	$GetPasswordInfoUser = Read-Host "Enter the UPN of the user you want to view the password last changed date for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password reset date for $GetPasswordInfoUser..."
		$textboxDetails.Text = "Get-MsolUser -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp"
		$TextboxResults.Text = Get-MsolUser -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting last password reset date for $GetPasswordInfoUser..."
		$textboxDetails.Text = "Get-MsolUser -TenantId $TenantText -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp"
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getPasswordLastResetDateForAllToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password reset date for all users..."
		$textboxDetails.Text = "Get-MsolUser | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, lastpasswordchangetimestamp -AutoSize "
		$TextboxResults.Text = Get-MsolUser | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, lastpasswordchangetimestamp -AutoSize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting last password reset date for all users..."
		$textboxDetails.Text = "Get-MsolUser -TenantId $TenantText | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, lastpasswordchangetimestamp -AutoSize "
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, lastpasswordchangetimestamp -AutoSize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$setPasswordToExpireForAUserToolStripMenuItem_Click = {
	$PasswordtoExpireforUser = Read-Host "Enter the UPN of the user you want the password to never expire for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to expire for $PasswordtoExpireforUser..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $PasswordtoExpireforUser –PasswordNeverExpires `$False"
		Set-MsolUser -UserPrincipalName $PasswordtoExpireforUser –PasswordNeverExpires $False
		$TextboxResults.text = Get-MSOLUser -UserPrincipalName $PasswordtoExpireforUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Setting password to expire for $PasswordtoExpireforUser..."
		$textboxDetails.Text = "Set-MsolUser -TenantId $TenantText -UserPrincipalName $PasswordtoExpireforUser –PasswordNeverExpires `$False"
		Set-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $PasswordtoExpireforUser –PasswordNeverExpires $False
		$TextboxResults.text = Get-MSOLUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $PasswordtoExpireforUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$setPasswordToNeverExpireForAUserToolStripMenuItem_Click = {
	$PasswordtoNeverExpireforUser = Read-Host "Enter the UPN of the user you want the password to expire for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to never expire for $PasswordtoNeverExpireforUser..."
		$textboxDetails.Text = "Set-MsolUser -UserPrincipalName $PasswordtoNeverExpireforUser –PasswordNeverExpires `$True"
		Set-MsolUser -UserPrincipalName $PasswordtoNeverExpireforUser –PasswordNeverExpires $True
		$TextboxResults.text = Get-MSOLUser -UserPrincipalName $PasswordtoNeverExpireforUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Setting password to never expire for $PasswordtoNeverExpireforUser..."
		$textboxDetails.Text = "Set-MsolUser -TenantId $TenantText -UserPrincipalName $PasswordtoNeverExpireforUser –PasswordNeverExpires `$True"
		Set-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $PasswordtoNeverExpireforUser –PasswordNeverExpires $True
		$TextboxResults.text = Get-MSOLUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $PasswordtoNeverExpireforUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getUsersWhosPasswordNeverExpiresToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting users where the password is set to never expire..."
		$textboxDetails.Text = "Get-MsolUser | Where-Object { `$_.PasswordNeverExpires -eq `$True } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires"
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.PasswordNeverExpires -eq $True } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting users where the password is set to never expire..."
		$textboxDetails.Text = "Get-MsolUser -TenantId $TenantText | Where-Object { `$_.PasswordNeverExpires -eq `$True } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires"
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.PasswordNeverExpires -eq $True } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getUsersWhosPasswordWillExpireToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting users where the password is set to expire..."
		$textboxDetails.Text = "Get-MsolUser | Where-Object { `$_.PasswordNeverExpires -eq `$False } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires"
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.PasswordNeverExpires -eq $False } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting users where the password is set to expire..."
		$textboxDetails.Text = "Get-MsolUser -TenantId $TenantText | Where-Object { `$_.PasswordNeverExpires -eq `$False } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires"
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.PasswordNeverExpires -eq $False } | Sort-Object UserPrincipalName | Format-Table UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getPasswordLastResetDateForAUserToolStripMenuItem_Click = {
	$GetPasswordInfoUser = Read-Host "Enter the UPN of the user you want to view the password last changed date for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password reset date for $GetPasswordInfoUser..."
		$textboxDetails.Text = "Get-MsolUser -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp"
		$TextboxResults.Text = Get-MsolUser -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting last password reset date for $GetPasswordInfoUser..."
		$textboxDetails.Text = "Get-MsolUser -TenantId $TenantText -UserPrincipalName $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp"
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getUsersNextPasswordResetDateToolStripMenuItem_Click = {
	$NextUserResetDateUser = Read-Host "Enter the UPN of the user"
	$VarDate = Read-Host "Enter days before passwords expires. EX: 90"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting $NextUserResetDateUser next password reset date..."
		$textboxDetails.Text = "(get-msoluser -userprincipalname $NextUserResetDateUser).lastpasswordchangetimestamp.adddays($VarDate) | Format-List DateTime"
		$TextboxResults.Text = (get-msoluser -userprincipalname $NextUserResetDateUser).lastpasswordchangetimestamp.adddays($VarDate) | Format-List DateTime | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting $NextUserResetDateUser next password reset date..."
		$textboxDetails.Text = "(get-msoluser -TenantId $TenantText -userprincipalname $NextUserResetDateUser).lastpasswordchangetimestamp.adddays($VarDate) | Format-List DateTime"
		$TextboxResults.Text = (get-msoluser -TenantId $PartnerComboBox.SelectedItem.TenantID -userprincipalname $NextUserResetDateUser).lastpasswordchangetimestamp.adddays($VarDate) | Format-List DateTime | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

	#Mailbox Permissions

$addFullPermissionsToAMailboxToolStripMenuItem_Click = {
	$mailboxAccess = read-host "Mailbox you want to give full-access to"
	$mailboxUser = read-host "Enter the UPN of the user that will have full access"
	try
	{
		$TextboxResults.Text = "Assigning full access permissions to $mailboxUser for the account $mailboxAccess..."
		$textboxDetails.Text = "Add-MailboxPermission $mailboxAccess -User $mailboxUser -AccessRights FullAccess -InheritanceType All"
		$TextboxResults.text = Add-MailboxPermission $mailboxAccess -User $mailboxUser -AccessRights FullAccess -InheritanceType All | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$addSendAsPermissionToAMailboxToolStripMenuItem_Click = {
	$SendAsAccess = read-host "Mailbox you want to give Send As access to"
	$mailboxUserAccess = read-host "Enter the UPN of the user that will have Send As access"
	try
	{
		$TextboxResults.Text = "Assigning Send-As access to $mailboxUserAccess for the account $SendAsAccess..."
		$textboxDetails.Text = "Add-RecipientPermission $SendAsAccess -Trustee $mailboxUserAccess -AccessRights SendAs"
		$TextboxResults.text = Add-RecipientPermission $SendAsAccess -Trustee $mailboxUserAccess -AccessRights SendAs | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$assignSendOnBehalfPermissionsForAMailboxToolStripMenuItem_Click = {
	$SendonBehalfof = read-host "Mailbox you want to give Send As access to"
	$mailboxUserSendonBehalfAccess = read-host "Enter the UPN of the user that will have Send As access"
	try
	{
		$TextboxResults.Text = "Assigning Send On Behalf of permissions to $mailboxUserSendonBehalfAccess for the account $SendonBehalfof..."
		$textboxDetails.Text = "Set-Mailbox -Identity $SendonBehalfof -GrantSendOnBehalfTo $mailboxUserSendonBehalfAccess"
		Set-Mailbox -Identity $SendonBehalfof -GrantSendOnBehalfTo $mailboxUserSendonBehalfAccess
		$TextboxResults.text = Get-Mailbox -Identity $SendonBehalfof | Format-List Identity, GrantSendOnBehalfTo | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$displayMailboxPermissionsForAUserToolStripMenuItem_Click = {
	$MailboxUserFullAccessPermission = Read-Host "Enter the UPN of the user want to view Full Access permissions for"
	try
	{
		$TextboxResults.Text = "Getting Full Access permissions for $MailboxUserFullAccessPermission..."
		$textboxDetails.Text = "Get-MailboxPermission $MailboxUserFullAccessPermission | Where-Object { (`$_.IsInherited -eq `$False) -and -not (`$_.User -like 'NT AUTHORITY\SELF') } | Sort-Object User | Format-Table -AutoSize"
		$TextboxResults.text = Get-MailboxPermission $MailboxUserFullAccessPermission | Where-Object { ($_.IsInherited -eq $False) -and -not ($_.User -like "NT AUTHORITY\SELF") } | Sort-Object User | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$displaySendAsPermissionForAMailboxToolStripMenuItem_Click = {
	$MailboxUserSendAsPermission = Read-Host "Enter the UPN of the user you want to view Send As permissions for"
	try
	{
		$TextboxResults.Text = "Getting Send As Permissions for $MailboxUserSendAsPermission..."
		$textboxDetails.Text = "Get-RecipientPermission $MailboxUserSendAsPermission | Where-Object { (`$_.IsInherited -eq `$False) -and -not (`$_.Trustee -like 'NT AUTHORITY\SELF') } | Sort-Object User | Format-Table -AutoSize"
		$TextboxResults.text = Get-RecipientPermission $MailboxUserSendAsPermission | Where-Object { ($_.IsInherited -eq $False) -and -not ($_.Trustee -like "NT AUTHORITY\SELF") } | Sort-Object User | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$displaySendOnBehalfPermissionsForMailboxToolStripMenuItem_Click = {
	$MailboxUserSendonPermission = Read-Host "Enter the UPN of the user you want to view Send On Behalf Of permission for"
	try
	{
		$TextboxResults.Text = "Getting Send On Behalf permissions for $MailboxUserSendonPermission..."
		$textboxDetails.Text = "Get-RecipientPermission $MailboxUserSendonPermission | Where-Object { (`$_.IsInherited -eq `$False) -and -not (`$_.Trustee -like 'NT AUTHORITY\SELF') } | Sort-Object User | Format-Table"
		$TextboxResults.text = Get-RecipientPermission $MailboxUserSendonPermission | Where-Object { ($_.IsInherited -eq $False) -and -not ($_.Trustee -like "NT AUTHORITY\SELF") } | Sort-Object User | Format-Table | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$removeFullAccessPermissionsForAMailboxToolStripMenuItem_Click = {
	$UserRemoveFullAccessRights = Read-Host "What user mailbox would you like modify Full Access rights to"
	$RemoveFullAccessRightsUser = Read-Host "Enter the UPN of the user you want to remove"
	try
	{
		$TextboxResults.Text = "Removing Full Access Permissions for $RemoveFullAccessRightsUser on account $UserRemoveFullAccessRights..."
		$textboxDetails.Text = "Remove-MailboxPermission  $UserRemoveFullAccessRights -User $RemoveFullAccessRightsUser -AccessRights FullAccess -Confirm:`$False -ea 1"
		Remove-MailboxPermission  $UserRemoveFullAccessRights -User $RemoveFullAccessRightsUser -AccessRights FullAccess -Confirm:$False -ea 1
		$TextboxResults.text = Get-MailboxPermission $UserRemoveFullAccessRights | Where-Object { ($_.IsInherited -eq $False) -and -not ($_.User -like "NT AUTHORITY\SELF") } | Sort-Object User | Format-Table | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$revokeSendAsPermissionsForAMailboxToolStripMenuItem_Click = {
	$UserDeleteSendAsAccessOn = Read-Host "What user mailbox would you like to modify Send As permission for?"
	$UserDeleteSendAsAccess = Read-Host "Enter the UPN of the user you want to remove Send As access to?"
	try
	{
		$TextboxResults.Text = "Removing Send As permission for $UserDeleteSendAsAccess on account $UserDeleteSendAsAccessOn..."
		$textboxDetails.Text = "Remove-RecipientPermission $UserDeleteSendAsAccessOn -AccessRights SendAs -Trustee $UserDeleteSendAsAccess"
		$TextboxResults.Text = Remove-RecipientPermission $UserDeleteSendAsAccessOn -AccessRights SendAs -Trustee $UserDeleteSendAsAccess | Sort-Object User | Format-Table | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$viewAllMailboxesAUserHasFullAccessToToolStripMenuItem_Click = {
	$ViewAllFullAccess = Read-Host "Enter the UPN of the account you want to view"
	try
	{
		$TextboxResults.Text = "Getting all mailboxes $ViewAllFullAccess has Full Access permissions to..."
		$textboxDetails.Text = "Get-Mailbox | Get-MailboxPermission -User $ViewAllFullAccess |  Sort-Object Identity | Format-Table"
		$TextboxResults.Text = Get-Mailbox | Get-MailboxPermission -User $ViewAllFullAccess | Sort-Object Identity | Format-Table | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$viewAllMailboxesAUserHasSendAsPermissionsToToolStripMenuItem_Click = {
	$ViewSendAsAccess = Read-Host "Enter the UPN of the account you want to view"
	try
	{
		$TextboxResults.Text = "Getting all mailboxes $ViewSendAsAccess has Send As permissions to..."
		$textboxDetails.Text = "Get-Mailbox | Get-RecipientPermission -Trustee $ViewSendAsAccess | Sort-Object Identity | Format-Table"
		$TextboxResults.Text = Get-Mailbox | Get-RecipientPermission -Trustee $ViewSendAsAccess | Sort-Object Identity | Format-Table | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$viewAllMailboxesAUserHasSendOnBehaldPermissionsToToolStripMenuItem_Click = {
	$ViewSendonBehalf = Read-Host "Enter the Name of the account you want to view"
	try
	{
		$TextboxResults.Text = "Getting all mailboxes $ViewSendonBehalf has Send On Behalf permissions to..."
		$textboxDetails.Text = "Get-Mailbox | Where-Object { `$_.GrantSendOnBehalfTo -match $ViewSendonBehalf } | Sort-Object DisplayName | Format-Table DisplayName, GrantSendOnBehalfTo, PrimarySmtpAddress, RecipientType"
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.GrantSendOnBehalfTo -match $ViewSendonBehalf } | Sort-Object DisplayName | Format-Table DisplayName, GrantSendOnBehalfTo, PrimarySmtpAddress, RecipientType | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$removeAllPermissionsToAMailboxToolStripMenuItem_Click = {
	$UserDeleteAllAccessOn = Read-Host "What user mailbox would you like to modify permissions for?"
	$UserDeleteAllAccess = Read-Host "Enter the UPN of the user you want to remove access to?"
	try
	{
		$TextboxResults.Text = "Removing all permissions for $UserDeleteAllAccess on account $UserDeleteAllAccessOn..."
		$textboxDetails.Text = "Remove-MailboxPermission -Identity $UserDeleteAllAccessOn -User $UserDeleteAllAccess -AccessRights FullAccess -InheritanceType All"
		$TextboxResults.Text = Remove-MailboxPermission -Identity $UserDeleteAllAccessOn -User $UserDeleteAllAccess -AccessRights FullAccess -InheritanceType All
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Forwarding

$getAllUsersForwardinToInternalRecipientToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all users forwarding to internal users..."
		$textboxDetails.Text = "Get-Mailbox | Where-Object { `$_.ForwardingAddress -ne `$Null -and `$_.RecipientType -eq 'UserMailbox' } | Sort-Object Name | Format-Table Name, ForwardingAddress, DeliverToMailboxAndForward -AutoSize"
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.ForwardingAddress -ne $Null -and $_.RecipientType -eq "UserMailbox" } | Sort-Object Name | Format-Table Name, ForwardingAddress, DeliverToMailboxAndForward -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$forwardToInternalRecipientAndDontSaveLocalCopyToolStripMenuItem_Click = {
	$UsertoFWD2 = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	$Fwd2me2 = Read-Host "Enter the Name, Display Name, Alias, or Email Address of the user to forward to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding from $UsertoFWD2 to $Fwd2me2..."
		$textboxDetails.Text = "Set-Mailbox  $UsertoFWD2 -ForwardingAddress $Fwd2me2 -DeliverToMailboxAndForward `$False"
		Set-Mailbox  $UsertoFWD2 -ForwardingAddress $Fwd2me2 -DeliverToMailboxAndForward $False
		$TextboxResults.Text = Get-Mailbox $UsertoFWD2 | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$forwardToExternalAddressAndSaveLocalCopyToolStripMenuItem_Click = {
	$UsertoFWD3 = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	$Fwd2me2External = Read-Host "Enter the external email address to forward to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding from $UsertoFWD3 to $Fwd2me2External..."
		$textboxDetails.Text = "Set-Mailbox $UsertoFWD3 -ForwardingsmtpAddress $Fwd2me2External -DeliverToMailboxAndForward `$true"
		Set-Mailbox $UsertoFWD3 -ForwardingsmtpAddress $Fwd2me2External -DeliverToMailboxAndForward $true
		$TextboxResults.Text = Get-Mailbox $UsertoFWD3 | Format-List Name, DeliverToMailboxAndForward, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$forwardToExternalAddressAndDontSaveLocalCopyToolStripMenuItem_Click = {
	$UsertoFWD4 = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	$Fwd2me2External2 = Read-Host "Enter the external email address to forward to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding from $UsertoFWD4 to $Fwd2me2External2..."
		$textboxDetails.Text = "Set-Mailbox $UsertoFWD4 -ForwardingsmtpAddress $Fwd2me2External2 -DeliverToMailboxAndForward `$False"
		Set-Mailbox $UsertoFWD4 -ForwardingsmtpAddress $Fwd2me2External2 -DeliverToMailboxAndForward $False
		$TextboxResults.Text = Get-Mailbox $UsertoFWD4 | Format-List Name, DeliverToMailboxAndForward, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getForwardingInfoForAUserToolStripMenuItem_Click = {
	$UserFwdInfo = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user"
	Try
	{
		$TextboxResults.Text = "Getting forwarding info for $UserFwdInfo..."
		$textboxDetails.Text = "Get-Mailbox $UserFwdInfo | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress"
		$TextboxResults.Text = Get-Mailbox $UserFwdInfo | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$removeExternalForwadingForAUserToolStripMenuItem_Click = {
	$RemoveFWDfromUserExternal = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address"
	Try
	{
		$TextboxResults.Text = "Removing all external forwarding from $RemoveFWDfromUserExternal..."
		$textboxDetails.Text = "Set-Mailbox $RemoveFWDfromUserExternal -ForwardingSmtpAddress `$Null"
		Set-Mailbox $RemoveFWDfromUserExternal -ForwardingSmtpAddress $Null
		$TextboxResults.Text = Get-Mailbox $RemoveFWDfromUserExternal | Format-List Name, DeliverToMailboxAndForward, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$removeAllForwardingForAUserToolStripMenuItem_Click = {
	$RemoveAllFWDforUser = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address"
	Try
	{
		$TextboxResults.Text = "Removing all forwarding from $RemoveAllFWDforUser..."
		$textboxDetails.Text = "Set-Mailbox $RemoveAllFWDforUser -ForwardingAddress `$Null -ForwardingSmtpAddress `$Null"
		Set-Mailbox $RemoveAllFWDforUser -ForwardingAddress $Null -ForwardingSmtpAddress $Null
		$TextboxResults.Text = Get-Mailbox $RemoveAllFWDforUser | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$removeInternalForwardingForUserToolStripMenuItem_Click = {
	$RemoveFWDfromUser = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address"
	Try
	{
		$TextboxResults.Text = "Removing all internal forwarding from $RemoveFWDfromUser..."
		$textboxDetails.Text = "Set-Mailbox $RemoveFWDfromUser -ForwardingAddress `$Null"
		Set-Mailbox $RemoveFWDfromUser -ForwardingAddress $Null
		$TextboxResults.Text = Get-Mailbox $RemoveFWDfromUser | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$forwardToInternalRecipientAndSaveLocalCopyToolStripMenuItem_Click = {
	$UsertoFWD = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	$Fwd2me = Read-Host "Enter the Name, Display Name, Alias, or Email Address of the user to forward to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding from $UsertoFWD to $Fwd2me..."
		$textboxDetails.Text = "Set-Mailbox  $UsertoFWD -ForwardingAddress $Fwd2me -DeliverToMailboxAndForward `$True"
		Set-Mailbox  $UsertoFWD -ForwardingAddress $Fwd2me -DeliverToMailboxAndForward $True
		$TextboxResults.Text = Get-Mailbox $UsertoFWD | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getAllUsersForwardingToExternalRecipientToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all users forwarding to internal users..."
		$textboxDetails.Text = "Get-Mailbox | Where-Object { `$_.ForwardingSmtpAddress -ne `$Null -and `$_.RecipientType -eq 'UserMailbox' } | Sort-Object Name | Format-Table Name, DeliverToMailboxAndForward, ForwardingSmtpAddress -AutoSize"
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.ForwardingSmtpAddress -ne $Null -and $_.RecipientType -eq "UserMailbox" } | Sort-Object Name |  Format-Table Name, DeliverToMailboxAndForward, ForwardingSmtpAddress -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
	
}

$removeAllForwardingForAllUsersToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Removing all forwarding from all users..."
		$textboxDetails.Text = "Get-Mailbox | Where-Object { `$_.RecipientType -eq 'UserMailbox' } | Set-Mailbox -ForwardingAddress `$Null -ForwardingSmtpAddress `$Null"
		$AllMailboxes = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingAddress $Null -ForwardingSmtpAddress $Null
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Sort-Object Name | Format-Table Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$removeExternalForwardingForAllUsersToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Removing all external forwarding from all users..."
		$textboxDetails.Text = "Get-Mailbox | Where-Object { `$_.RecipientType -eq 'UserMailbox' } | Set-Mailbox -ForwardingSmtpAddress `$Null"
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingSmtpAddress $Null
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Sort-Object Name | Format-Table Name, DeliverToMailboxAndForward, ForwardingSmtpAddress -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$removeInternalForwardingForAllUsersToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Removing all internal forwarding from all users..."
		$textboxDetails.Text = "Get-Mailbox | Where-Object { `$_.RecipientType -eq 'UserMailbox' } | Set-Mailbox -ForwardingAddress `$Null"
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingAddress $Null
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Sort-Object Name | Format-Table Name, DeliverToMailboxAndForward, ForwardingAddress -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$forwardAllUsersEmailToExternalRecipientAndSaveALocalCopyToolStripMenuItem_Click = {
	$ForwardAllToExternal = Read-Host "Enter the email to forward all email to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding for all users to $ForwardAllToExternal..."
		$textboxDetails.Text = "Get-Mailbox | Where-Object { `$_.RecipientType -eq 'UserMailbox' } | Set-Mailbox -ForwardingsmtpAddress $ForwardAllToExternal -DeliverToMailboxAndForward `$True"
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingsmtpAddress $ForwardAllToExternal -DeliverToMailboxAndForward $True
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Sort-Object Name | Format-Table Name, DeliverToMailboxAndForward, ForwardingSmtpAddress -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$forwardAllUsersEmailToExternalRecipientAndDontSaveALocalCopyToolStripMenuItem_Click = {
	$ForwardAllToExternal2 = Read-Host "Enter the email to forward all email to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding for all users to $ForwardAllToExternal2..."
		$textboxDetails.Text = "Get-Mailbox | Where-Object { `$_.RecipientType -eq 'UserMailbox' } | Set-Mailbox -ForwardingsmtpAddress $ForwardAllToExternal2 -DeliverToMailboxAndForward `$False"
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingsmtpAddress $ForwardAllToExternal2 -DeliverToMailboxAndForward $False
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Sort-Object Name | Format-Table Name, DeliverToMailboxAndForward, ForwardingSmtpAddress -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$forwardAllUsersEmailToInternalRecipientAndSaveLocalCopyToolStripMenuItem_Click = {
	$ForwardAllToInternal = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding for all users to $ForwardAllToInternal..."
		$textboxDetails.Text = "Get-Mailbox | Where-Object { `$_.RecipientType -eq 'UserMailbox' } | Set-Mailbox -ForwardingAddress $ForwardAllToInternal -DeliverToMailboxAndForward `$True"
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingAddress $ForwardAllToInternal -DeliverToMailboxAndForward $True
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Sort-Object Name | Format-Table Name, DeliverToMailboxAndForward, ForwardingAddress | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$forwardAllUsersEmailToInternalRecipientAndDontSaveLocalCopyToolStripMenuItem_Click = {
	$ForwardAllToInternal2 = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding for all users to $ForwardAllToInternal2..."
		$textboxDetails.Text = "Get-Mailbox | Where-Object { `$_.RecipientType -eq 'UserMailbox' } | Set-Mailbox -ForwardingAddress $ForwardAllToInternal2 -DeliverToMailboxAndForward `$True"
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingAddress $ForwardAllToInternal2 -DeliverToMailboxAndForward $True
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Sort-Object Name | Format-Table Name, DeliverToMailboxAndForward, ForwardingAddress | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}



###GROUPS###

	#Distribution Groups

$displayDistributionGroupsToolStripMenuItem_Click={
	try
	{
		$TextboxResults.Text = "Getting all Distribution Groups..."
		$textboxDetails.Text = "Get-DistributionGroup | Where-Object { `$_.GroupType -notlike 'Universal, SecurityEnabled'} | Sort-Object DisplayName | Format-Table DisplayName -AutoSize"
		$TextboxResults.text = Get-DistributionGroup | Where-Object { $_.GroupType -notlike "Universal, SecurityEnabled"} | Sort-Object DisplayName | Format-Table DisplayName -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$createADistributionGroupToolStripMenuItem_Click = {
	$NewDistroGroup = Read-Host "What is the name of the new Distribution Group?"
	try
	{
		$TextboxResults.Text = "Creating the $NewDistroGroup Distribution Group..."
		$textboxDetails.Text = "New-DistributionGroup -Name $NewDistroGroup | Format-List"
		$TextboxResults.Text = New-DistributionGroup -Name $NewDistroGroup | Format-List | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$deleteADistributionGroupToolStripMenuItem_Click = {
	$DeleteDistroGroup = Read-Host "Enter the name of the Distribtuion group you want deleted."
	try
	{
		$TextboxResults.Text = "Deleting the $DeleteDistroGroup Distribution Group..."
		$textboxDetails.Text = "Remove-DistributionGroup $DeleteDistroGroup"
		Remove-DistributionGroup $DeleteDistroGroup
		$TextboxResults.Text = "Getting list of distribution groups"
		$TextboxResults.text = Get-DistributionGroup | Where-Object { $_.GroupType -notlike "Universal, SecurityEnabled" } | Sort-Object DisplayName | Format-Table DisplayName | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$allowDistributionGroupToReceiveExternalEmailToolStripMenuItem_Click = {
	$AllowExternalEmail = Read-Host "Enter the name of the Distribtuion Group you want to allow external email to"
	try
	{
		$TextboxResults.Text = "Allowing extneral senders for the $AllowExternalEmail Distribution Group..."
		$textboxDetails.Text = "Set-DistributionGroup $AllowExternalEmail -RequireSenderAuthenticationEnabled `$False"
		Set-DistributionGroup $AllowExternalEmail -RequireSenderAuthenticationEnabled $False 
		$TextboxResults.text = Get-DistributionGroup $AllowExternalEmail | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$hideDistributionGroupFromGALToolStripMenuItem_Click = {
	$GroupHideGAL = Read-Host "Enter the name of the Distribtuion Group you want to allow external email to"
	try
	{
		$TextboxResults.Text = "Hiding the $GroupHideGAL from the Global Address List..."
		$textboxDetails.Text = "Set-DistributionGroup $GroupHideGAL -HiddenFromAddressListsEnabled `$True"
		Set-DistributionGroup $GroupHideGAL -HiddenFromAddressListsEnabled $True
		$TextboxResults.text = Get-DistributionGroup $GroupHideGAL | Format-List Name, HiddenFromAddressListsEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$displayDistributionGroupMembersToolStripMenuItem_Click = {
	$ListDistributionGroupMembers = Read-Host "Enter the name of the Distribution Group you want to list members of"
	try
	{
		$TextboxResults.Text = "Getting all members of the $ListDistributionGroupMembers Distrubution Group..."
		$textboxDetails.Text = "Get-DistributionGroupMember $ListDistributionGroupMembers | Sort-Object DisplayName | Format-Table -AutoSize"
		$TextboxResults.Text = Get-DistributionGroupMember $ListDistributionGroupMembers | Sort-Object DisplayName | Format-Table -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$addAUserToADistributionGroupToolStripMenuItem_Click = {
	$DistroGroupAdd = Read-Host "Enter the name of the Distribution Group"
	$DistroGroupAddUser = Read-Host "Enter the UPN of the user you wish to add to $DistroGroupAdd"
	try
	{
		$TextboxResults.Text = "Adding $DistroGroupAddUser to the $DistroGroupAdd Distribution Group..."
		$textboxDetails.Text = "Add-DistributionGroupMember -Identity $DistroGroupAdd -Member $DistroGroupAddUser"
		Add-DistributionGroupMember -Identity $DistroGroupAdd -Member $DistroGroupAddUser
		$TextboxResults.Text = "Getting members of $DistroGroupAdd..."
		$TextboxResults.Text = Get-DistributionGroupMember $DistroGroupAdd | Sort-Object DisplayName | Format-Table -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$removeAUserADistributionGroupToolStripMenuItem_Click = {
	$DistroGroupRemove = Read-Host "Enter the name of the Distribution Group"
	$DistroGroupRemoveUser = Read-Host "Enter the UPN of the user you wish to remove from $DistroGroupRemove"
	try
	{
		$TextboxResults.Text = "Removing $DistroGroupRemoveUser from the $DistroGroupRemove Distribution Group..."
		$textboxDetails.Text = "Remove-DistributionGroupMember -Identity $DistroGroupRemove -Member $DistroGroupRemoveUser"
		Remove-DistributionGroupMember -Identity $DistroGroupRemove -Member $DistroGroupRemoveUser
		$TextboxResults.Text = "Getting members of $DistroGroupRemove"
		$TextboxResults.Text = Get-DistributionGroupMember $DistroGroupRemove | Sort-Object DisplayName | Format-Table -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$addAllUsersToADistributionGroupToolStripMenuItem_Click = {
		$users = Get-Mailbox | Select-Object -ExpandProperty Alias
		$AddAllUsersToSingleDistro = Read-Host "Enter the name of the Distribution Group you want to add all users to"
		try
		{
		$TextboxResults.Text = "Adding all users to the $AddAllUsersToSingleDistro distribution group..."
		$textboxDetails.Text = "Foreach (`$user in `$users) { Add-DistributionGroupMember -Identity $AddAllUsersToSingleDistro -Member `$user }"
		Foreach ($user in $users) { Add-DistributionGroupMember -Identity $AddAllUsersToSingleDistro -Member $user }
		$TextboxResults.Text = "Getting members of $AddAllUsersToSingleDistro"
		$TextboxResults.Text = Get-DistributionGroupMember $AddAllUsersToSingleDistro | Sort-Object DisplayName | Format-Table -AutoSize | Out-String
		}
		catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
}

$getDetailedInfoForDistributionGroupToolStripMenuItem_Click = {
	$DetailedInfoMailDistroGroup = Read-Host "Enter the group name"
	Try
	{
		$TextboxResults.Text = "Getting detailed info about the $DetailedInfoMailDistroGroup group..."
		$textboxDetails.Text = "Get-DistributionGroup $DetailedInfoMailDistroGroup | Format-List"
		$TextboxResults.text = Get-DistributionGroup $DetailedInfoMailDistroGroup | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$allowAllDistributionGroupsToReceiveExternalEmailToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Allowing extneral senders for all Distribution Groups..."
		$textboxDetails.Text = "Get-DistributionGroup | Set-DistributionGroup -RequireSenderAuthenticationEnabled `$False"
		Get-DistributionGroup | Set-DistributionGroup -RequireSenderAuthenticationEnabled $False
		$TextboxResults.text = Get-DistributionGroup | Sort-Object Name | Format-Table Name, RequireSenderAuthenticationEnabled -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$denyDistributionGroupFromReceivingExternalEmailToolStripMenuItem_Click = {
	$DenyExternalEmail = Read-Host "Enter the name of the Distribtuion Group you want to deny external email to"
	try
	{
		$TextboxResults.Text = "Denying extneral senders for the $DenyExternalEmail Distribution Group..."
		$textboxDetails.Text = "Set-DistributionGroup $DenyExternalEmail -RequireSenderAuthenticationEnabled `$True"
		Set-DistributionGroup $DenyExternalEmail -RequireSenderAuthenticationEnabled $True
		$TextboxResults.text = Get-DistributionGroup $DenyExternalEmail | Sort-Object Name | Format-Table Name, RequireSenderAuthenticationEnabled -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$denyAllDistributionGroupsFromReceivingExternalEmailToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Denying extneral senders for all Distribution Groups..."
		$textboxDetails.Text = "Get-DistributionGroup | Set-DistributionGroup -RequireSenderAuthenticationEnabled `$True"
		Get-DistributionGroup | Set-DistributionGroup -RequireSenderAuthenticationEnabled $True
		$TextboxResults.text = Get-DistributionGroup | Sort-Objects Name | Format-Table Name, RequireSenderAuthenticationEnabled -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getExternalEmailStatusForADistributionGroupToolStripMenuItem_Click = {
	$ExternalEmailStatus = Read-Host "Enter the Distribution Group"
	try
	{
		$TextboxResults.Text = "Getting external email status for $ExternalEmailStatus..."
		$textboxDetails.Text = "Get-DistributionGroup $ExternalEmailStatus | Format-List Name, RequireSenderAuthenticationEnabled"
		$TextboxResults.text = Get-DistributionGroup $ExternalEmailStatus | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getExternalEmailStatusForAllDistributionGroupsToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting external email status for all distribution groups..."
		$textboxDetails.Text = "Get-DistributionGroup | Sort-Object Name | Format-Table Name, RequireSenderAuthenticationEnabled -AutoSize"
		$TextboxResults.text = Get-DistributionGroup | Sort-Object Name | Format-Table Name, RequireSenderAuthenticationEnabled -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

	#Unified Groups

$getListOfUnifiedGroupsToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting list of all unified groups..."
		$textboxDetails.Text = "Get-UnifiedGroup | Sort-Object Name | Format-Table -AutoSize"
		$TextboxResults.Text = Get-UnifiedGroup | Sort-Object Name | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$listMembersOfAGroupToolStripMenuItem_Click = {
	$GetUnifiedGroupMembers = Read-Host "Enter the name of the group you want to view members for."
	try
	{
		$TextboxResults.Text = "Getting all members of the $GetUnifiedGroupMembers group..."
		$textboxDetails.Text = "Get-UnifiedGroupLinks –Identity $GetUnifiedGroupMembers –LinkType Members | Sort-Object DisplayName | Format-Table DisplayName -AutoSize"
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $GetUnifiedGroupMembers –LinkType Members | Sort-Object DisplayName | Format-Table DisplayName -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$removeAGroupToolStripMenuItem_Click = {
	$RemoveUnifiedGroup = Read-Host "Enter the name of the group you want to remove"
	try
	{
		$TextboxResults.Text = "Removing the $RemoveUnifiedGroup group..."
		$textboxDetails.Text = "Remove-UnifiedGroup $RemoveUnifiedGroup"
		Remove-UnifiedGroup $RemoveUnifiedGroup
		$TextboxResults.Text = "Getting list of unified groups..."
		$TextboxResults.Text = Get-UnifiedGroup | Sort-Object Name | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$addAUserToAGroupToolStripMenuItem_Click = {
	$UnifiedGroupAddUser = Read-Host "Enter the name of the group you want to add a user to"
	$UnifiedGroupUserAdd = Read-Host "Enter the UPN of the user you want to add to the $UnifiedGroupAddUser group."
	$TextboxResults.Text = "Access Levels:
Members
Owners
Subscribers"
	$UnifiedGroupAccess = Read-Host "Enter the level of access you want $UnifiedGroupUserAdd to have."
	try
	{
		$TextboxResults.Text = "Adding $UnifiedGroupUserAdd as a member of the $UnifiedGroupAddUser group..."
		$textboxDetails.Text = "Add-UnifiedGroupLinks $UnifiedGroupAddUser –Links $UnifiedGroupUserAdd –LinkType $UnifiedGroupAccess"
		Add-UnifiedGroupLinks $UnifiedGroupAddUser –Links $UnifiedGroupUserAdd –LinkType $UnifiedGroupAccess
		$TextboxResults.Text = "Getting members for $UnifiedGroupAddUser..."
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $UnifiedGroupAddUser –LinkType Members | Sort-Object DisplayName | Format-Table DisplayName -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$createANewGroupToolStripMenuItem_Click = {
	$NewUnifiedGroupName = Read-Host "Enter the Display Name of the new group"
	$NewUnifiedGroupAccessType = Read-Host "Enter the Access Type for the group $NewUnifiedGroupName (Public or Private)"
	try
	{
		$TextboxResults.Text = "Creating a the $NewUnifiedGroupName group..."
		$textboxDetails.Text = "New-UnifiedGroup –DisplayName $NewUnifiedGroupName -AccessType $NewUnifiedGroupAccessType"
		New-UnifiedGroup –DisplayName $NewUnifiedGroupName -AccessType $NewUnifiedGroupAccessType
		$TextboxResults.Text = Get-UnifiedGroup $NewUnifiedGroupName | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$listOwnersOfAGroupToolStripMenuItem_Click = {
	$GetUnifiedGroupOwners = Read-Host "Enter the name of the group you want to view members for."
	try
	{
		$TextboxResults.Text = "Getting all owners of the $GetUnifiedGroupOwners group..."
		$textboxDetails.Text = "Get-UnifiedGroupLinks –Identity $GetUnifiedGroupOwners –LinkType Owners | Format-List DisplayName | Format-Table -AutoSize"
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $GetUnifiedGroupOwners –LinkType Owners | Sort-Object DisplayName | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$listSubscribersOfAGroupToolStripMenuItem_Click = {
	$GetUnifiedGroupSubscribers = Read-Host "Enter the name of the group you want to view members for."
	try
	{
		$TextboxResults.Text = "Getting all subscribers of the $GetUnifiedGroupSubscribers group..."
		$textboxDetails.Text = "Get-UnifiedGroupLinks –Identity $GetUnifiedGroupSubscribers –LinkType Subscribers | Sort-Object DisplayName | Format-Table -AutoSize"
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $GetUnifiedGroupSubscribers –LinkType Subscribers | Sort-Object DisplayName | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$addAnOwnerToAGroupToolStripMenuItem_Click = {
	$TextboxResults.Text = "Important! The user must be a member of the group prior to becoming an owner"
	$UnifiedGroupAddOwner = Read-Host "Enter the name of the group you want to modify ownership for"
	$AddUnifiedGroupOwner = Read-Host "Enter the UPN of the user you want to become an owner"
	try
	{
		$TextboxResults.Text = "Adding $AddUnifiedGroupOwner as an owner of the $UnifiedGroupAddOwner group..."
		$textboxDetails.Text = "Add-UnifiedGroupLinks -Identity $UnifiedGroupAddOwner -LinkType Owners -Links $AddUnifiedGroupOwner"
		Add-UnifiedGroupLinks -Identity $UnifiedGroupAddOwner -LinkType Owners -Links $AddUnifiedGroupOwner
		$TextboxResults.Text = "Getting list of owners for $UnifiedGroupAddOwner..."
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $UnifiedGroupAddOwner –LinkType Owners | Sort-Object DisplayName | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$addASubscriberToAGroupToolStripMenuItem_Click = {
	$UnifiedGroupAddSubscriber = Read-Host "Enter the name of the group you want to add a subscriber for"
	$AddUnifiedGroupSubscriber = Read-Host "Enter the UPN of the user you want to add as a subscriber"
	try
	{
		$TextboxResults.Text = "Adding $AddUnifiedGroupSubscriber as a subscriber to the $UnifiedGroupAddSubscriber group..."
		$textboxDetails.Text = "Add-UnifiedGroupLinks -Identity $UnifiedGroupAddSubscriber -LinkType Owners -Links $AddUnifiedGroupSubscriber"
		Add-UnifiedGroupLinks -Identity $UnifiedGroupAddSubscriber -LinkType Owners -Links $AddUnifiedGroupSubscriber
		$TextboxResults.Text = "Getting list of subscribers for $UnifiedGroupAddSubscriber..."
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $UnifiedGroupAddSubscriber –LinkType Subscribers | Sort-Object DisplayName | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

	#Security Groups

$createARegularSecurityGroupToolStripMenuItem_Click = {
	$SecurityGroupName = Read-Host "Enter a name for the new Security Group"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Creating the $SecurityGroupName security group..."
		$textboxDetails.Text = "New-MsolGroup -DisplayName $SecurityGroupName | Format-List | Out-String"
		$TextboxResults.Text = New-MsolGroup -DisplayName $SecurityGroupName | Format-List | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Creating the $SecurityGroupName security group..."
		$textboxDetails.Text = "New-MsolGroup -DisplayName $SecurityGroupName -TenantId $TenantText"
		$TextboxResults.Text = New-MsolGroup -DisplayName $SecurityGroupName -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getAllRegularSecurityGroupsToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting list of all Security groups..."
		$textboxDetails.Text = "Get-MsolGroup -GroupType Security | Sort-Object DisplayName | Format-Table DisplayName, GroupType -AutoSize"
		$TextboxResults.Text = Get-MsolGroup -GroupType Security | Sort-Object DisplayName | Format-Table DisplayName, GroupType -AutoSize | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting list of all Security groups..."
		$textboxDetails.Text = "Get-MsolGroup -TenantId $TenantText -GroupType Security | Sort-Object DisplayName | Format-Table DisplayName, GroupType -AutoSize"
		$TextboxResults.Text = Get-MsolGroup -TenantId $PartnerComboBox.SelectedItem.TenantID -GroupType Security | Sort-Object DisplayName | Format-Table DisplayName, GroupType -AutoSize | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getAllMailEnabledSecurityGroupsToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all Mail Enabled Security Groups..."
		$textboxDetails.Text = "Get-DistributionGroup | Where-Object { `$_.GroupType -like 'Universal, SecurityEnabled' } | Sort-Object DisplayName | Format-Table DisplayName, SamAccountName, GroupType, IsDirSynced, EmailAddresses -AutoSize "
		$TextboxResults.text = Get-DistributionGroup | Where-Object { $_.GroupType -like "Universal, SecurityEnabled" } | Sort-Object DisplayName | Format-Table DisplayName, SamAccountName, GroupType, IsDirSynced, EmailAddresses -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$createAMailEnabledSecurityGroupToolStripMenuItem_Click = {
	$MailEnabledSecurityGroup = Read-Host "Enter the name of the security group"
	$MailEnabledSMTPAddress = Read-Host "Enter the primary SMTP address for $MailEnabledSecurityGroup"
	Try
	{
		$TextboxResults.Text = "Creating the $MailEnabledSecurityGroup security group..."
		$textboxDetails.Text = "New-DistributionGroup -Name $MailEnabledSecurityGroup -Type Security -PrimarySmtpAddress $MailEnabledSMTPAddress"
		$TextboxResults.Text = New-DistributionGroup -Name $MailEnabledSecurityGroup -Type Security -PrimarySmtpAddress $MailEnabledSMTPAddress | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$addAUserToAMailEnabledSecurityGroupToolStripMenuItem_Click = {
	$MailEnabledGroupAdd = Read-Host "Enter the name of the Group"
	$MailEnabledGroupAddUser = Read-Host "Enter the UPN of the user you wish to add to $MailEnabledGroupAdd"
	try
	{
		$TextboxResults.Text = "Adding $MailEnabledGroupAddUser to the $MailEnabledGroupAdd Group..."
		$textboxDetails.Text = "Add-DistributionGroupMember -Identity $MailEnabledGroupAdd -Member $MailEnabledGroupAddUser"
		Add-DistributionGroupMember -Identity $MailEnabledGroupAdd -Member $MailEnabledGroupAddUser
		$TextboxResults.Text = "Getting members of $MailEnabledGroupAdd..."
		$TextboxResults.Text = Get-DistributionGroupMember $MailEnabledGroupAdd | Sort-Object DisplayName | Format-Table Displayname -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$allowSecurityGroupToRecieveExternalMailToolStripMenuItem_Click = {
	$AllowExternalEmailSecurity = Read-Host "Enter the name of the Group you want to allow external email to"
	try
	{
		$TextboxResults.Text = "Allowing extneral senders for the $AllowExternalEmailSecurity Group..."
		$textboxDetails.Text = "Set-DistributionGroup $AllowExternalEmailSecurity -RequireSenderAuthenticationEnabled `$False"
		Set-DistributionGroup $AllowExternalEmailSecurity -RequireSenderAuthenticationEnabled $False
		$TextboxResults.text = Get-DistributionGroup $AllowExternalEmailSecurity | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getDetailedInfoForMailEnabledSecurityGroupToolStripMenuItem_Click = {
	$DetailedInfoMailEnabledSecurityGroup = Read-Host "Enter the group name"
	Try
	{
		$TextboxResults.Text = "Getting detailed info about the $DetailedInfoMailEnabledSecurityGroup group..."
		$textboxDetails.Text = "Get-DistributionGroup $DetailedInfoMailEnabledSecurityGroup | Format-List"
		$TextboxResults.text = Get-DistributionGroup $DetailedInfoMailEnabledSecurityGroup | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$removeMailENabledSecurityGroupToolStripMenuItem_Click = {
	$DeleteMailEnabledSecurityGroup = Read-Host "Enter the name of the group you want deleted."
	try
	{
		$TextboxResults.Text = "Deleting the $DeleteMailEnabledSecurityGroup Group..."
		$textboxDetails.Text = "Remove-DistributionGroup $DeleteMailEnabledSecurityGroup"
		Remove-DistributionGroup $DeleteMailEnabledSecurityGroup
		$TextboxResults.Text = "Getting list of mail enabled security groups..."
		$TextboxResults.text = Get-DistributionGroup | Where-Object { $_.GroupType -like "Universal, SecurityEnabled" } | Sort-Object DisplayName | Format-Table DisplayName | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$denySecurityGroupFromRecievingExternalEmailToolStripMenuItem_Click = {
	$DenyExternalEmailSecurity = Read-Host "Enter the name of the Group you want to deny external email to"
	try
	{
		$TextboxResults.Text = "Denying extneral senders for the $DenyExternalEmailSecurity Group..."
		$textboxDetails.Text = "Set-DistributionGroup $DenyExternalEmailSecurity -RequireSenderAuthenticationEnabled `$True"
		Set-DistributionGroup $DenyExternalEmailSecurity -RequireSenderAuthenticationEnabled $True
		$TextboxResults.text = Get-DistributionGroup $DenyExternalEmailSecurity | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}



###RESOURCE MAILBOX###

	#Booking Options

$allowConflictMeetingsToolStripMenuItem_Click = {
	$ConflictMeetingAllow = Read-Host "Enter the Room Name of the Resource Calendar you want to allow conflicts"
	try
	{
		$TextboxResults.Text = "Allowing conflict meetings $ConflictMeetingAllow..."
		$textboxDetails.Text = "Set-CalendarProcessing $ConflictMeetingAllow -AllowConflicts `$True"
		Set-CalendarProcessing $ConflictMeetingAllow -AllowConflicts $True
		$TextboxResults.Text = Get-CalendarProcessing -identity $ConflictMeetingAllow | Format-List Identity, AllowConflicts | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$denyConflictMeetingsForAllResourceMailboxesToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Denying conflict meeting for all room calendars..."
		$textboxDetails.Text = "Get-MailBox | Where-Object { `$_.ResourceType -eq 'Room' } | Set-CalendarProcessing -AllowConflicts `$False"
		Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AllowConflicts $False
		$TextboxResults.Text = Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Get-CalendarProcessing | Sort-Object Identity | Format-Table Identity, AllowConflicts -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$allowConflicMeetingsForAllResourceMailboxesToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Allowing conflict meeting for all room calendars..."
		$textboxDetails.Text = "Get-MailBox | Where-Object { `$_.ResourceType -eq 'Room' } | Set-CalendarProcessing -AllowConflicts `$True"
		Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AllowConflicts $True
		$TextboxResults.Text = Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Get-CalendarProcessing | Sort-Object Identity | Format-Table Identity, AllowConflicts -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$disallowconflictmeetingsToolStripMenuItem_Click = {
	$ConflictMeetingDeny = Read-Host "Enter the Room Name of the Resource Calendar you want to disallow conflicts"
	try
	{
		$TextboxResults.Text = "Denying conflict meetings for $ConflictMeetingDeny..."
		$textboxDetails.Text = "Set-CalendarProcessing $ConflictMeetingDeny -AllowConflicts `$False"
		Set-CalendarProcessing $ConflictMeetingDeny -AllowConflicts $False
		$TextboxResults.Text = Get-CalendarProcessing -identity $ConflictMeetingDeny | Format-List Identity, AllowConflicts | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$enableAutomaticBookingForAllResourceMailboxToolStripMenuItem1_Click = {
		Try
		{
		$TextboxResults.Text = "Enabling automatic booking on all room calendars..."
		$textboxDetails.Text = "Get-MailBox | Where-Object { `$_.ResourceType -eq 'Room' } | Set-CalendarProcessing -AutomateProcessing:AutoAccept"
		Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AutomateProcessing:AutoAccept
		$TextboxResults.Text = Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Get-CalendarProcessing | Sort-Object Identity | Format-Table Identity, AutomateProcessing -AutoSize | Out-String
		}
		Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
	
}

$GetRoomMailBoxCalendarProcessingToolStripMenuItem_Click = {
	$RoomMailboxCalProcessing = Read-Host "Enter the Calendar Name you want to view calendar processing information for"
	try
	{
		$TextboxResults.Text = "Getting calendar processing information for $RoomMailboxCalProcessing..."
		$textboxDetails.Text = "Get-Mailbox $RoomMailboxCalProcessing | Get-CalendarProcessing | Format-List"
		$TextboxResults.Text = Get-Mailbox $RoomMailboxCalProcessing | Get-CalendarProcessing | Format-List | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

	#Room Mailbox

$convertAMailboxToARoomMailboxToolStripMenuItem_Click = {
	$MailboxtoRoom = Read-Host "What user would you like to convert to a Room Mailbox? Please enter the full email address"
	Try
	{
		$TextboxResults.Text = "Converting $MailboxtoRoom to a Room Mailbox..."
		$textboxDetails.Text = "Set-Mailbox $MailboxtoRoom -Type Room"
		Set-Mailbox $MailboxtoRoom -Type Room
		$TextboxResults.Text = Get-MailBox $MailboxtoRoom | Format-List Name, ResourceType, PrimarySmtpAddress, EmailAddresses, UserPrincipalName, IsMailboxEnabled | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$createANewRoomMailboxToolStripMenuItem_Click = {
	$NewRoomMailbox = Read-Host "Enter the name of the new room mailbox"
	Try
	{
		$TextboxResults.Text = "Creating the $NewRoomMailbox Room Mailbox..."
		$textboxDetails.Text = "New-Mailbox -Name $NewRoomMailbox -Room "
		$TextboxResults.Text = New-Mailbox -Name $NewRoomMailbox -Room | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getListOfRoomMailboxesToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting list of all Room Mailboxes..."
		$textboxDetails.Text = "Get-MailBox | Where-Object { `$_.ResourceType -eq 'Room' } | Sort-Object Name | Format-Table -AutoSize"
		$TextboxResults.Text = Get-MailBox | Where-Object { $_.ResourceType -eq "Room" }  | Sort-Object Name |  Format-Table -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$removeARoomMailboxToolStripMenuItem_Click = {
	$RemoveRoomMailbox = Read-Host "Enter the name of the room mailbox"
	Try
	{
		$TextboxResults.Text = "Removing the $RemoveRoomMailbox Room Mailbox..."
		$textboxDetails.Text = "Remove-Mailbox $RemoveRoomMailbox"
		Remove-Mailbox $RemoveRoomMailbox
		$TextboxResults.Text = "Getting list of Room Mailboxes..."
		$TextboxResults.Text = Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Sort-Object Name | Format-Table -AutoSize | Out-String
		
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}




###JUNK EMAIL CONFIGURATION###

	#Blacklist

$blacklistDomainForAllToolStripMenuItem_Click = {
	$BlacklistDomain = Read-Host "Enter the domain you want to blacklist for all users. EX: google.com"
	try
	{
		$TextboxResults.Text = "Blacklisting $BlacklistDomain for all users..."
		$textboxDetails.Text = "Get-Mailbox | Set-MailboxJunkEmailConfiguration -BlockedSendersAndDomains `@{ Add = $BlacklistDomain }"
		Get-Mailbox | Set-MailboxJunkEmailConfiguration -BlockedSendersAndDomains @{ Add = $BlacklistDomain } 
		$TextboxResults.Text = Get-Mailbox | Get-MailboxJunkEmailConfiguration | Sort-Object Identity | Format-Table Identity, BlockedSendersAndDomains, Enabled -AutoSize | Out-String
		
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$blacklistDomainForASingleUserToolStripMenuItem_Click = {
	$Blockeddomainuser = Read-Host "Enter the UPN of the user you want to modify junk email for"
	$BlockedDomain2 = Read-Host "Enter the domain you want to blacklist"
	try
	{
		$TextboxResults.Text = "Blacklisting $BlockedDomain2 for $Blockeddomainuser..."
		$textboxDetails.Text = "Set-MailboxJunkEmailConfiguration -Identity $Blockeddomainuser -BlockedSendersAndDomains `@{ Add = $BlockedDomain2 } "
		Set-MailboxJunkEmailConfiguration -Identity $Blockeddomainuser -BlockedSendersAndDomains @{ Add = $BlockedDomain2 } 
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $Blockeddomainuser | Format-List Identity, BlockedSendersAndDomains | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$blacklistASpecificEmailAddressForAllToolStripMenuItem_Click = {
	$BlockSpecificEmailForAll = Read-Host "Enter the email address you want to blacklist for all"
	try
	{
		$TextboxResults.Text = "Blacklisting $BlockSpecificEmailForAll for all users..."
		$textboxDetails.Text = "Get-Mailbox | Set-MailboxJunkEmailConfiguration -BlockedSendersAndDomains `@{ Add = $BlockSpecificEmailForAll }"
		Get-Mailbox | Set-MailboxJunkEmailConfiguration -BlockedSendersAndDomains @{ Add = $BlockSpecificEmailForAll }
		$TextboxResults.Text = Get-Mailbox | Get-MailboxJunkEmailConfiguration | Sort-Object Identity | Format-Table Identity, BlockedSendersAndDomains, Enabled -Autosize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$blacklistASpecificEmailAddressForASingleUserToolStripMenuItem_Click = {
	$ModifyblacklistforaUser = Read-Host "Enter the user you want to modify the blacklist for"
	$DenySpecificEmailForOne = Read-Host "Enter the email address you want to whitelist for a single user"
	try
	{
		$TextboxResults.Text = "Blacklisting $DenySpecificEmailForOne for $ModifyblacklistforaUser..."
		$textboxDetails.Text = "Set-MailboxJunkEmailConfiguration -Identity $ModifyblacklistforaUser -BlockedSendersAndDomains `@{ Add = $DenySpecificEmailForOne }"
		Set-MailboxJunkEmailConfiguration -Identity $ModifyblacklistforaUser -BlockedSendersAndDomains @{ Add = $DenySpecificEmailForOne }
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $ModifyblacklistforaUser | Format-List Identity, BlockedSendersAndDomains, Enabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Junk Email General Items

$checkSafeAndBlockedSendersForAUserToolStripMenuItem_Click = {
	$CheckSpamUser = Read-Host "Enter the UPN of the user you want to check blocked and allowed senders for"
	try
	{
		$TextboxResults.Text = "Getting safe and blocked senders for $CheckSpamUser..."
		$textboxDetails.Text = "Get-MailboxJunkEmailConfiguration -Identity $CheckSpamUser | Format-List Identity, TrustedListsOnly, ContactsTrusted, TrustedSendersAndDomains, BlockedSendersAndDomains, TrustedRecipientsAndDomains, IsValid "
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $CheckSpamUser | Format-List Identity, TrustedListsOnly, ContactsTrusted, TrustedSendersAndDomains, BlockedSendersAndDomains, TrustedRecipientsAndDomains, IsValid | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Whitelist

$whitelistDomainForAllToolStripMenuItem_Click = {
	$AllowedDomain = Read-Host "Enter the domain you want to whitelist for all users. EX: google.com"
	try
	{
		$TextboxResults.Text = "Whitelisting $AllowedDomain for all..."
		$textboxDetails.Text = "Get-Mailbox | Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains `@{ Add = $AllowedDomain }"
		Get-Mailbox | Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains @{ Add = $AllowedDomain }
		$TextboxResults.Text = Get-Mailbox | Get-MailboxJunkEmailConfiguration | Sort-Object Identity | Format-Table Identity, TrustedSendersAndDomains, TrustedRecipientsAndDomains -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$whitelistDomainForASingleUserToolStripMenuItem_Click = {
	$Alloweddomainuser = Read-Host "Enter the UPN of the user you want to modify junk email for"
	$AllowedDomain2 = Read-Host "Enter the domain you want to whitelist"
	try
	{
		$TextboxResults.Text = "Whitelisting $AllowedDomain2 for $Alloweddomainuser..."
		$textboxDetails.Text = "Set-MailboxJunkEmailConfiguration -Identity $Alloweddomainuser -TrustedSendersAndDomains `@{ Add = $AllowedDomain2 }"
		Set-MailboxJunkEmailConfiguration -Identity $Alloweddomainuser -TrustedSendersAndDomains @{ Add = $AllowedDomain2 }
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $Alloweddomainuser | Format-List Identity, TrustedSendersAndDomains, TrustedRecipientsAndDomains | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	} 
}

$whitelistASpecificEmailAddressForAllToolStripMenuItem_Click = {
	$AllowSpecificEmailForAll = Read-Host "Enter the email address you want to whitelist for all"
	try
	{
		$TextboxResults.Text = "Whitelisting $AllowSpecificEmailForAll for all..."
		$textboxDetails.Text = "Get-Mailbox | | Where-Object { `$_.RecipientType -eq 'UserMailbox' } | Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains @{ Add = $AllowSpecificEmailForAll }"
		Get-Mailbox | Where-Object { $_.RecipientType -eq 'UserMailbox' } |  Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains @{ Add = $AllowSpecificEmailForAll }
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq 'UserMailbox' } | Get-MailboxJunkEmailConfiguration | Sort-Object Identity | Format-Table Identity, TrustedSendersAndDomains, TrustedRecipientsAndDomains -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$whitelistASpecificEmailAddressForASingleUserToolStripMenuItem_Click = {
	$ModifyWhitelistforaUser = Read-Host "Enter the user you want to modify the whitelist for"
	$AllowSpecificEmailForOne = Read-Host "Enter the email address you want to whitelist for a single user"
	try
	{
		$TextboxResults.Text = "Whitelisting $AllowSpecificEmailForOne for $ModifyWhitelistforaUser..."
		$textboxDetails.Text = "Set-MailboxJunkEmailConfiguration -Identity $ModifyWhitelistforaUser -TrustedSendersAndDomains `@{ Add = $AllowSpecificEmailForOne }"
		Set-MailboxJunkEmailConfiguration -Identity $ModifyWhitelistforaUser -TrustedSendersAndDomains @{ Add = $AllowSpecificEmailForOne }
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $ModifyWhitelistforaUser | Format-List Identity, TrustedSendersAndDomains, TrustedRecipientsAndDomains, Enabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}


###CONTACTS FOLDER PERMISSIONS

$addContactsPermissionsForAUserToolStripMenuItem_Click = {
	$ContacsUser = Read-Host "Enter the UPN of the user whose contacts folder you want to give access to"
	$Contactsuser2 = Read-Host "Enter the UPN of the user who you want to give access to"
	$TextboxResults.text = "Contacts Permissions: 
Owner
PublishingEditor
Editor
PublishingAuthor
Author
NonEditingAuthor
Reviewer
Contributor
AvailabilityOnly
LimitedDetails"
	$level = Read-Host "Access Level?"
	try
	{
		$TextboxResults.Text = "Adding $Contactsuser2 to $ContacsUser contacts folder with $level permissions..."
		$textboxDetails.Text = "Add-MailboxFolderPermission -Identity ${Calendaruser}:\contacts -user $Calendaruser2 -AccessRights $level"
		Remove-MailboxFolderPermission -identity ${ContacsUser}:\contacts -user $Contactsuser2 -Confirm:$False
		Add-MailboxFolderPermission -Identity ${ContacsUser}:\contacts -user $Contactsuser2 -AccessRights $level
		$TextboxResults.Text = "Getting contact folder permissions for $ContacsUser..."
		$TextboxResults.Text = Get-MailboxFolderPermission -Identity ${ContacsUser}:\contacts | Sort-Object User, AccessRights | Format-Table User, AccessRights, FolderName -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getUsersContactsPermissionsToolStripMenuItem_Click = {
	$ContactsUserPermissions = Read-Host "What user would you like contacts folder permissions for?"
	Try
	{
		$TextboxResults.Text = "Getting $ContactsUserPermissions contacts permissions..."
		$textboxDetails.Text = "Get-MailboxFolderPermission -Identity ${ContactsUserPermissions}:\contacts | Sort-Object User, AccessRights | Format-Table User, AccessRights, FolderName "
		$TextboxResults.Text = Get-MailboxFolderPermission -Identity ${ContactsUserPermissions}:\contacts | Sort-Object User, AccessRights | Format-Table User, AccessRights, FolderName -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$removeAUserFromSomeonesContactsPermissionsToolStripMenuItem_Click = {
	$Contactsuserremove = Read-Host "Enter the UPN of the user whose contacts you want to remove access to"
	$Contactsuser2remove = Read-Host "Enter the UPN of the user who you want to remove access"
	try
	{
		$TextboxResults.Text = "Removing $Contactsuser2remove from $Contactsuserremove contacts folder..."
		$textboxDetails.Text = "Remove-MailboxFolderPermission -Identity ${Contactsuserremove}:\contacts -user $Contactsuser2remove"
		Remove-MailboxFolderPermission -Identity ${Contactsuserremove}:\contacts -user $Calendaruser2remove
		$TextboxResults.Text = "Getting contact folder permissions for $Contactsuserremove..."
		$TextboxResults.Text = Get-MailboxFolderPermission -Identity ${Contactsuserremove}:\contacts | Sort-Object User, AccessRights | Format-Table User, AccessRights, FolderName -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$removeAUserFromAllContactsFoldersToolStripMenuItem_Click = {
	$RemoveUserFromAllContacts = Read-Host "Enter the UPN of the user you want to remove from all contacts folders"
	try
	{
		$TextboxResults.Text = "Removing $RemoveUserFromAllContacts from all users contacts folders..."
		$textboxDetails.Text = "`$users = Get-Mailbox | Select-Object -ExpandProperty Alias
Foreach (`$user in `$users) { Remove-MailboxFolderPermission `${user}:\Contacts -user $RemoveUserFromAllContacts -Confirm:`$false}﻿"
		$users = Get-Mailbox | Select-Object -ExpandProperty Alias
		Foreach ($user in $users) { Remove-MailboxFolderPermission ${user}:\Contacts -user $RemoveUserFromAllContacts -Confirm:$false }﻿
	}
	catch
	{
		#[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	$TextboxResults.Text = ""
	
}

$addASingleUserPermissionsOnAllContactsFoldersToolStripMenuItem_Click = {
	$MasterUserContacts = Read-Host "Enter the UPN of the user you want permission on all users contacts folders"
	$TextboxResults.text = "Contacts Permissions: 
Owner
PublishingEditor
Editor
PublishingAuthor
Author
NonEditingAuthor
Reviewer
Contributor
AvailabilityOnly
LimitedDetails"
	$level2 = Read-Host "Access Level?"
	try
	{
		$TextboxResults.Text = "Adding $MasterUserContacts to everyones contacts folder with $level2 permissions..."
		$textboxDetails.Text = "Get-Mailbox | Select-Object -ExpandProperty Alias
Foreach (`$user in `$users) { Add-MailboxFolderPermission `${user}:\Contacts -user $MasterUserContacts -accessrights $level2 }﻿"
		$users = Get-Mailbox | Select-Object -ExpandProperty Alias
		Foreach ($user in $users) { Add-MailboxFolderPermission ${user}:\Contacts -user $MasterUserContacts -accessrights $level2 }﻿
	}
	catch
	{
		#[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	$TextboxResults.Text = ""
	
}



###ADMIN###

	#OWA

$disableAccessToOWAForASingleUserToolStripMenuItem_Click = {
	$DisableOWAforUser = Read-Host "Enter the UPN of the user you want to disable OWA access for"
	try
	{
		$TextboxResults.Text = "Disabling OWA access for $DisableOWAforUser..."
		$textboxDetails.Text = "Set-CASMailbox $DisableOWAforUser -OWAEnabled `$False"
		Set-CASMailbox $DisableOWAforUser -OWAEnabled $False
		$TextboxResults.Text = Get-CASMailbox $DisableOWAforUser | Format-List DisplayName, OWAEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$enableOWAAccessForASingleUserToolStripMenuItem_Click = {
	$EnableOWAforUser = Read-Host "Enter the UPN of the user you want to enable OWA access for"
	try
	{
		$TextboxResults.Text = "Enabling OWA access for $EnableOWAforUser..."
		$textboxDetails.Text = "Set-CASMailbox $EnableOWAforUser -OWAEnabled `$True"
		Set-CASMailbox $EnableOWAforUser -OWAEnabled $True
		$TextboxResults.Text = Get-CASMailbox $EnableOWAforUser | Format-List DisplayName, OWAEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$disableOWAAccessForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Disabling OWA access for all..."
		$textboxDetails.Text = "Get-Mailbox | Set-CASMailbox -OWAEnabled `$False"
		Get-Mailbox | Set-CASMailbox -OWAEnabled $False
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox | Sort-Object DisplayName | Format-Table DisplayName, OWAEnabled -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$enableOWAAccessForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Enabling OWA access for all..."
		$textboxDetails.Text = "Get-Mailbox | Set-CASMailbox -OWAEnabled `$True"
		Get-Mailbox | Set-CASMailbox -OWAEnabled $True
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox | Sort-Object DisplayName | Format-Table DisplayName, OWAEnabled -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getOWAAccessForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting OWA info for all users..."
		$textboxDetails.Text = "Get-Mailbox | Get-CASMailbox | Sort-Object DisplayName | Format-Table DisplayName, OWAEnabled, OwaMailboxPolicy, OWAforDevicesEnabled -Autosize"
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox | Sort-Object DisplayName | Format-Table DisplayName, OWAEnabled, OwaMailboxPolicy, OWAforDevicesEnabled -Autosize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getOWAInfoForASingleUserToolStripMenuItem_Click = {
	$OWAAccessUser = Read-Host "Enter the UPN for the User you want to view OWA info for"
	try
	{
		$TextboxResults.Text = "Getting OWA Access for $OWAAccessUser..."
		$textboxDetails.Text = "Get-CASMailbox -identity $OWAAccessUser | Format-List DisplayName, OWAEnabled, OwaMailboxPolicy, OWAforDevicesEnabled"
		$TextboxResults.Text = Get-CASMailbox -identity $OWAAccessUser | Format-List DisplayName, OWAEnabled, OwaMailboxPolicy, OWAforDevicesEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#ActiveSync

$getActiveSyncDevicesForAUserToolStripMenuItem_Click = {
	$ActiveSyncDevicesUser = Read-Host "Enter the UPN of the user you wish to see ActiveSync Devices for"
	try
	{
		$TextboxResults.Text = "Getting ActiveSync device info for $ActiveSyncDevicesUser..."
		$textboxDetails.Text = "Get-MobileDeviceStatistics -Mailbox $ActiveSyncDevicesUser | Format-List DeviceFriendlyName, DeviceModel, DeviceOS, DeviceMobileOperator, DeviceType, Status, FirstSyncTime, LastPolicyUpdateTime, LastSyncAttemptTime, LastSuccessSync, LastPingHeartbeat, DeviceAccessState, IsValid "
		$TextboxResults.Text = Get-MobileDeviceStatistics -Mailbox $ActiveSyncDevicesUser | Format-List DeviceFriendlyName, DeviceModel, DeviceOS, DeviceMobileOperator, DeviceType, Status, FirstSyncTime, LastPolicyUpdateTime, LastSyncAttemptTime, LastSuccessSync, LastPingHeartbeat, DeviceAccessState, IsValid  | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$disableActiveSyncForAUserToolStripMenuItem_Click = {
	$DisableActiveSyncForUser = Read-Host "Enter the UPN of the user you wish to disable ActiveSync for"
	try
	{
		$TextboxResults.Text = "Disabling ActiveSync for $DisableActiveSyncForUser..."
		$textboxDetails.Text = "Set-CASMailbox $DisableActiveSyncForUser -ActiveSyncEnabled `$False"
		Set-CASMailbox $DisableActiveSyncForUser -ActiveSyncEnabled $False 
		$TextboxResults.Text = Get-CASMailbox -Identity $DisableActiveSyncForUser | Format-List DisplayName, ActiveSyncEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$enableActiveSyncForAUserToolStripMenuItem_Click = {
	$EnableActiveSyncForUser = Read-Host "Enter the UPN of the user you wish to enable ActiveSync for"
	try
	{
		$TextboxResults.Text = "Enabling ActiveSync for $EnableActiveSyncForUser... "
		$textboxDetails.Text = "Set-CASMailbox $EnableActiveSyncForUser -ActiveSyncEnabled `$True"
		Set-CASMailbox $EnableActiveSyncForUser -ActiveSyncEnabled $True
		$TextboxResults.Text = Get-CASMailbox -Identity $EnableActiveSyncForUser | Format-List DisplayName, ActiveSyncEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$viewActiveSyncInfoForAUserToolStripMenuItem_Click = {
	$ActiveSyncInfoForUser = Read-Host "Enter the UPN for the user you want to view ActiveSync info for"
	try
	{
		$TextboxResults.Text = "Getting ActiveSync info for $ActiveSyncInfoForUser..."
		$textboxDetails.Text = "Get-CASMailbox -Identity $ActiveSyncInfoForUser | Format-List DisplayName, ActiveSyncEnabled, ActiveSyncAllowedDeviceIDs, ActiveSyncBlockedDeviceIDs, ActiveSyncMailboxPolicy, ActiveSyncMailboxPolicyIsDefaulted, ActiveSyncDebugLogging, HasActiveSyncDevicePartnership"
		$TextboxResults.Text = Get-CASMailbox -Identity $ActiveSyncInfoForUser | Format-List DisplayName, ActiveSyncEnabled, ActiveSyncAllowedDeviceIDs, ActiveSyncBlockedDeviceIDs, ActiveSyncMailboxPolicy, ActiveSyncMailboxPolicyIsDefaulted, ActiveSyncDebugLogging, HasActiveSyncDevicePartnership | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$disableActiveSyncForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Disabling ActiveSync for all..."
		$textboxDetails.Text = "Get-Mailbox | Set-CASMailbox -ActiveSyncEnabled `$False"
		Get-Mailbox | Set-CASMailbox -ActiveSyncEnabled $False
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox  | Sort-Object DisplayName | Format-Table DisplayName, ActiveSyncEnabled -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getActiveSyncInfoForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting ActiveSync info for all..."
		$textboxDetails.Text = "Get-Mailbox | Get-CASMailbox | Sort-Object DisplayName | Format-Table DisplayName, ActiveSyncEnabled -AutoSize"
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox | Sort-Object DisplayName | Format-Table DisplayName, ActiveSyncEnabled -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		
	}
}

$enableActiveSyncForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Enabling ActiveSync for all.."
		$textboxDetails.Text = "Get-Mailbox | Set-CASMailbox -ActiveSyncEnabled `$True"
		Get-Mailbox | Set-CASMailbox -ActiveSyncEnabled $True
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox | Sort-Object DisplayName | Format-Table DisplayName, ActiveSyncEnabled -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#PowerShell

$disableAccessToPowerShellForAUserToolStripMenuItem_Click = {
	$DisablePowerShellforUser = Read-Host "Enter the UPN of the user you want to disable PowerShell access for"
	try
	{
		$TextboxResults.Text = "Disabling PowerShell access for $DisablePowerShellforUser..."
		$textboxDetails.Text = "Set-User $DisablePowerShellforUser -RemotePowerShellEnabled `$False"
		Set-User $DisablePowerShellforUser -RemotePowerShellEnabled $False
		$TextboxResults.Text = Get-User $DisablePowerShellforUser | Format-List DisplayName, RemotePowerShellEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$displayPowerShellRemotingStatusForAUserToolStripMenuItem_Click = {
	$PowerShellRemotingStatusUser = Read-Host "Enter the UPN of the user you want to view PowerShell Remoting for"
	try
	{
		$TextboxResults.Text = "Getting PowerShell info for $PowerShellRemotingStatusUser..."
		$textboxDetails.Text = "Get-User $PowerShellRemotingStatusUser | Format-List DisplayName, RemotePowerShellEnabled"
		$TextboxResults.Text = Get-User $PowerShellRemotingStatusUser | Format-List DisplayName, RemotePowerShellEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$enableAccessToPowerShellForAUserToolStripMenuItem_Click = {
	$EnablePowerShellforUser = Read-Host "Enter the UPN of the user you want to enable PowerShell access for"
	try
	{
		$TextboxResults.Text = "Enabling PowerShell access for $EnablePowerShellforUser..."
		$textboxDetails.Text = "Set-User $EnablePowerShellforUser -RemotePowerShellEnabled `$True"
		Set-User $EnablePowerShellforUser -RemotePowerShellEnabled $True
		$TextboxResults.Text = Get-User $EnablePowerShellforUser | Format-List DisplayName, RemotePowerShellEnabled | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getPowerShellRemotingStatusForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting PowerShell info for all users..."
		$textboxDetails.Text = "Get-User | Sort-Object DisplayName | Format-Table DisplayName, RemotePowerShellEnabled -AutoSize"
		$TextboxResults.Text = Get-User | Sort-Object DisplayName | Format-Table DisplayName, RemotePowerShellEnabled -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}


#Message Trace

$messageTraceToolStripMenuItem_Click = {
	
}

$GetAllRecentToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting recent messages..."
		$textboxDetails.Text = "Get-MessageTrace | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status"
		$TextboxResults.Text = Get-MessageTrace | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$fromACertainSenderToolStripMenuItem1_Click = {
	$MessageTraceSender = Read-Host "Enter the senders email address"
	try
	{
		$TextboxResults.Text = "Getting messages from $MessageTraceSender..."
		$textboxDetails.Text = "Get-MessageTrace -SenderAddress $MessageTraceSender | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status"
		$TextboxResults.Text = Get-MessageTrace -SenderAddress $MessageTraceSender | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$toACertainRecipientToolStripMenuItem_Click = {
	$MessageTraceRecipient = Read-Host "Enter the recipients email address"
	try
	{
		$TextboxResults.Text = "Getting messages sent to $MessageTraceRecipient..."
		$textboxDetails.Text = "Get-MessageTrace -RecipientAddress $MessageTraceRecipient | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status"
		$TextboxResults.Text = Get-MessageTrace -RecipientAddress $MessageTraceRecipient | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getFailedMessagesToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting failed messages..."
		$textboxDetails.Text = "Get-MessageTrace -Status 'Failed' | Format-Table Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status"
		$TextboxResults.Text = Get-MessageTrace -Status "Failed" | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$GetMessagesBetweenDatesToolStripMenuItem_Click = {
	$MessageTraceStartDate = Read-Host "Enter the start date. EX: 06/13/2015 or 09/01/2015 5:00 PM"
	$MessageTraceEndDate = Read-Host "Enter the end date. EX: 06/15/2015 or 09/01/2015 5:00 PM"
	try
	{
		$TextboxResults.Text = "Getting messages between $MessageTraceStartDate and $MessageTraceEndDate..."
		$textboxDetails.Text = "Get-MessageTrace -StartDate $MessageTraceStartDate -EndDate $MessageTraceEndDate | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status"
		$TextboxResults.Text = Get-MessageTrace -StartDate $MessageTraceStartDate -EndDate $MessageTraceEndDate | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Company Info

$getTechnicalNotificationEmailToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting technical account info..."
		$textboxDetails.Text = "Get-MsolCompanyInformation | Format-List TechnicalNotificationEmails"
		$TextboxResults.Text = Get-MsolCompanyInformation | Format-List TechnicalNotificationEmails | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting technical account info..."
		$textboxDetails.Text = "Get-MsolCompanyInformation -TenantId $TenantText | Format-List TechnicalNotificationEmails"
		$TextboxResults.Text = Get-MsolCompanyInformation -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List TechnicalNotificationEmails | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$lastDirSyncTimeToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last DirSync time..."
		$textboxDetails.Text = "Get-MsolCompanyInformation | Format-List LastDirSyncTime"
		$TextboxResults.Text = Get-MsolCompanyInformation | Format-List LastDirSyncTime | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting last DirSync time..."
		$textboxDetails.Text = "Get-MsolCompanyInformation -TenantId $TenantText | Format-List LastDirSyncTime"
		$TextboxResults.Text = Get-MsolCompanyInformation -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List LastDirSyncTime | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getLastPasswordSyncTimeToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password sync time..."
		$textboxDetails.Text = "Get-MsolCompanyInformation | Format-List LastPasswordSyncTime"
		$TextboxResults.Text = Get-MsolCompanyInformation | Format-List LastPasswordSyncTime | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting last password sync time..."
		$textboxDetails.Text = "Get-MsolCompanyInformation -TenantId $TenantText  | Format-List LastPasswordSyncTime"
		$TextboxResults.Text = Get-MsolCompanyInformation -TenantId $PartnerComboBox.SelectedItem.TenantID  | Format-List LastPasswordSyncTime | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getAllCompanyInfoToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all company info..."
		$textboxDetails.Text = "Get-MsolCompanyInformation | Format-List "
		$TextboxResults.Text = Get-MsolCompanyInformation | Format-List | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TenantText = $PartnerComboBox.text
		$TextboxResults.Text = "Getting all company info..."
		$textboxDetails.Text = "Get-MsolCompanyInformation -TenantId $TenantText | Format-List"
		$TextboxResults.Text = Get-MsolCompanyInformation -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List | Out-String
	}
	Else
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Sharing Policy

$getSharingPolicyToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting all sharing policies..."
		$textboxDetails.Text = "Get-SharingPolicy | Format-List Name, Domains, Enabled, Default, Identity, WhenChanged, WhenCreated, IsValid, ObjectState"
		$TextboxResults.Text = Get-SharingPolicy | Format-List Name, Domains, Enabled, Default, Identity, WhenChanged, WhenCreated, IsValid, ObjectState | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$disableASharingPolicyToolStripMenuItem_Click = {
	$DisableSharingPolicy = Read-Host "Enter the name of the policy you want to disable"
	try
	{
		$TextboxResults.Text = "Disabling the sharing policy $DisableSharingPolicy..."
		$textboxDetails.Text = "Set-SharingPolicy -Identity $DisableSharingPolicy -Enabled `$False"
		Set-SharingPolicy -Identity $DisableSharingPolicy -Enabled $False
		$TextboxResults.Text = Get-SharingPolicy -Identity $DisableSharingPolicy | Format-List Name, Enabled, ObjectState | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$enableASharingPolicyToolStripMenuItem_Click = {
	$EnableSharingPolicy = Read-Host "Enter the name of the policy you want to Enable"
	try
	{
		$TextboxResults.Text = "Enabling the sharing policy $EnableSharingPolicy..."
		$textboxDetails.Text = "Set-SharingPolicy -Identity $EnableSharingPolicy -Enabled `$True"
		Set-SharingPolicy -Identity $EnableSharingPolicy -Enabled $True
		$TextboxResults.Text = Get-SharingPolicy -Identity $EnableSharingPolicy | Format-List Name, Enabled, ObjectState | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$createANewSharingPolicyToolStripMenuItem_Click = {
	$TextboxResults.Text = "You may need to enable organization customization if customization status is dehydrated."
	$TextboxResults.Text = Get-OrganizationConfig | Format-List  Identity, IsDehydrated | Out-String
	$NewSharingPolicyName = Read-Host "Enter the name for the sharing policy"
	$TextboxResults.Text = "The following sharing policy action values can be used:
CalendarSharingFreeBusySimple: Share free/busy hours only
CalendarSharingFreeBusyDetail: Share free/busy hours, subject, and location
CalendarSharingFreeBusyReviewer: Share free/busy hours, subject, location, and the body of the message or calendar item
ContactsSharing: Share contacts only

EXAMPLE: mail.contoso.com: CalendarSharingFreeBusyDetail, ContactsSharing "
	$SharingPolicy = Read-Host "Enter the domain this policy will apply to and the value"
	try
	{
		$TextboxResults.Text = "Creating a new sharing policy $NewSharingPolicyName..."
		$textboxDetails.Text = "New-SharingPolicy -Name $NewSharingPolicyName -Domains $SharingPolicy"
		New-SharingPolicy -Name $NewSharingPolicyName -Domains $SharingPolicy
		$TextboxResults.Text = Get-SharingPolicy -Identity $NewSharingPolicyName | Format-List Name, Domains, Enabled, Default, Identity, WhenChanged, WhenCreated, IsValid, ObjectState | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getInfoForASingleSharingPolicyToolStripMenuItem_Click = {
	$DetailedInfoForSharingPolicy = Read-Host "Enter the name of the policy you want info for"
	try
	{
		$TextboxResults.Text = "Getting info for the sharing policy $DetailedInfoForSharingPolicy..."
		$textboxDetails.Text = "Get-SharingPolicy -Identity $DetailedInfoForSharingPolicy | Format-List Name, Domains, Enabled, Default, Identity, WhenChanged, WhenCreated, IsValid, ObjectState"
		$TextboxResults.Text = Get-SharingPolicy -Identity $DetailedInfoForSharingPolicy | Format-List Name, Domains, Enabled, Default, Identity, WhenChanged, WhenCreated, IsValid, ObjectState | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Configuration 

$enableCustomizationToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Enabling customization..."
		$textboxDetails.Text = "Enable-OrganizationCustomization"
		Enable-OrganizationCustomization
		$TextboxResults.Text = "Success"
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getCustomizationStatusToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting customization status..."
		$textboxDetails.Text = "Get-OrganizationConfig | Format-Table  Identity, IsDehydrated -AutoSize "
		$TextboxResults.Text = Get-OrganizationConfig | Format-Table  Identity, IsDehydrated -AutoSize | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getOrganizationCustomizationToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting customization information..."
		$textboxDetails.Text = "Get-OrganizationConfig | Format-List"
		$TextboxResults.Text = Get-OrganizationConfig | Format-List | Out-String
	}
	catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getSharepointSiteToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting sharepoint URL"
		$textboxDetails.Text = "Get-OrganizationConfig | Format-List SharePointUrl"
		$TextboxResults.Text = Get-OrganizationConfig | Format-List SharePointUrl | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}



###Reporting###

$getAllMailboxSizesToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Generating mailbox sizes report..."
		$textboxDetails.Text = "Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics | Select-Object DisplayName,`@{name = 'TotalItemSize (MB)'; expression = {[math]::Round( `
(`$_.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(', ', '')/1MB), 2)}}, `
ItemCount | Sort-Object 'TotalItemSize (MB)' -Descending | Format-Table -AutoSize"
		$TextboxResults.Text =
		Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics | Select-Object DisplayName, `
		
		@{
			name = "TotalItemSize (MB)"; expression = {
				[math]::Round( `
				($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "")/1MB), 2)
			}
		}, `
		
		ItemCount | Sort-Object "TotalItemSize (MB)" -Descending | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Mail Malware Report

$getMailMalwareReportToolStripMenuItem1_Click = {
	$TextboxResults.Text = "Generating recent mail malware report..."
	Try
	{
		$TextboxResults.Text = "Getting mail malware report..."
		$textboxDetails.Text = "Get-MailDetailMalwareReport | Format-List"
		$TextboxResults.Text = Get-MailDetailMalwareReport | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getMailMalwareReportFromSenderToolStripMenuItem_Click = {
	$MalwareSender = Read-Host "Enter the email of the sender"
	try
	{
		$TextboxResults.Text = "Generating mail malware report sent from $MalwareSender..."
		$textboxDetails.Text = "Get-MailDetailMalwareReport -SenderAddress $MalwareSender | Format-List"
		$TextboxResults.Text = Get-MailDetailMalwareReport -SenderAddress $MalwareSender | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getMailMalwareReportBetweenDatesToolStripMenuItem_Click = {
	$MalwareReportStart = Read-Host "Enter the start date. EX: 06/13/2015"
	$MalwareReportEnd = Read-Host "Enter the end date. EX: 06/15/2015 "
	try
	{
		$TextboxResults.Text = "Generating mail malware report between $MalwareReportStart and $MalwareReportEnd..."
		$textboxDetails.Text = "Get-MailDetailMalwareReport -StartDate $MalwareReportStart -EndDate $MalwareReportEnd | Format-List"
		$TextboxResults.Text = Get-MailDetailMalwareReport -StartDate $MalwareReportStart -EndDate $MalwareReportEnd | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getMailMalwareReportToARecipientToolStripMenuItem_Click = {
	$MalwareRecipient = Read-Host "Enter the email of the recipient"
	try
	{
		$TextboxResults.Text = "Generating mail malware report sent to $MalwareRecipient..."
		$textboxDetails.Text = "Get-MailDetailMalwareReport -RecipientAddress $MalwareRecipient | Format-List"
		$TextboxResults.Text = Get-MailDetailMalwareReport -RecipientAddress $MalwareRecipient | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getMailMalwareReporforInboundToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Generating mail malware inbound report..."
		$textboxDetails.Text = "Get-MailDetailMalwareReport -Direction Inbound | Format-List "
		$TextboxResults.Text = Get-MailDetailMalwareReport -Direction Inbound | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getMailMalwareReportForOutboundToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Generating mail malware outbound report..."
		$textboxDetails.Text = "Get-MailDetailMalwareReport -Direction Outbound | Format-List"
		$TextboxResults.Text = Get-MailDetailMalwareReport -Direction Outbound | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Mail Traffic Report

$getRecentMailTrafficReportToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Generating recent mail traffic report..."
		$textboxDetails.Text = "Get-MailTrafficReport | Format-Table -AutoSize"
		$TextboxResults.Text = Get-MailTrafficReport | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getInboundMailTrafficReportToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Generating inbound traffic report..."
		$textboxDetails.Text = "Get-MailTrafficReport -Direction Inbound | Format-Table -AutoSize"
		$TextboxResults.Text = Get-MailTrafficReport -Direction Inbound | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getOutboundMailTrafficReportToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Generating outbound mail traffic report..."
		$textboxDetails.Text = "Get-MailTrafficReport -Direction Outbound | Format-Table -AutoSize"
		$TextboxResults.Text = Get-MailTrafficReport -Direction Outbound | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getMailTrafficReportBetweenDatesToolStripMenuItem_Click = {
	$MailTrafficStart = Read-Host "Enter the start date. EX: 06/13/2015"
	$MailTrafficEnd = Read-Host "Enter the end date. EX: 06/15/2015 "
	Try
	{
		$TextboxResults.Text = "Generating mail traffic report between $MailTrafficStart and $MailTrafficEnd..."
		$textboxDetails.Text = "Get-MailTrafficReport -StartDate $MailTrafficStart -EndDate $MailTrafficEnd | Format-Table -AutoSize"
		$TextboxResults.Text = Get-MailTrafficReport -StartDate $MailTrafficStart -EndDate $MailTrafficEnd | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}



###SHARED MAILBOXES###

$createASharedMailboxToolStripMenuItem_Click = {
	$NewSharedMailbox = Read-Host "Enter the name of the new Shared Mailbox"
	Try
	{
		$TextboxResults.Text = "Creating new shared mailbox $NewSharedMailbox"
		$textboxDetails.Text = "New-Mailbox -Name $NewSharedMailbox –Shared"
		New-Mailbox -Name $NewSharedMailbox –Shared
		$TextboxResults.Text = Get-Mailbox -RecipientTypeDetails SharedMailbox | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getAllSharedMailboxesToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting list of shared mailboxes..."
		$textboxDetails.Text = "Get-Mailbox -RecipientTypeDetails SharedMailbox | Format-Table -AutoSize"
		$TextboxResults.Text = Get-Mailbox -RecipientTypeDetails SharedMailbox | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$convertRegularMailboxToSharedToolStripMenuItem_Click = {
	$ConvertMailboxtoShared = Read-Host "Enter the name of the account you want to convert"
	Try
	{
		$TextboxResults.Text = "Converting $ConvertMailboxtoShared to a shared mailbox..."
		$textboxDetails.Text = "Set-Mailbox $ConvertMailboxtoShared –Type shared"
		Set-Mailbox $ConvertMailboxtoShared –Type shared
		$TextboxResults.Text = Get-Mailbox -Identity $ConvertMailboxtoShared | Format-List UserPrincipalName, DisplayName, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses, IsDirSynced | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$convertSharedMailboxToRegularToolStripMenuItem_Click = {
	$ConvertMailboxtoRegular = Read-Host "Enter the name of the account you want to convert"
	Try
	{
		$TextboxResults.Text = "Converting $ConvertMailboxtoRegular to a regular mailbox..."
		$textboxDetails.Text = "Set-Mailbox $ConvertMailboxtoRegular –Type Regular"
		Set-Mailbox $ConvertMailboxtoRegular –Type Regular
		$TextboxResults.Text = Get-Mailbox -Identity $ConvertMailboxtoRegular | Format-List UserPrincipalName, DisplayName, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses, IsDirSynced | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getDetailedInfoForASharedMailboxToolStripMenuItem_Click = {
	$SharedMailboxDetailedInfo = Read-Host "Enter the name of the shared mailbox"
	Try
	{
		$TextboxResults.Text = "Getting shared mailbox information for $SharedMailboxDetailedInfo..."
		$textboxDetails.Text = "Get-Mailbox $SharedMailboxDetailedInfo | Format-List"
		$TextboxResults.Text = Get-Mailbox $SharedMailboxDetailedInfo | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

	#Permissions

$exportAllUsersMailboxPermissionsToCSVToolStripMenuItem_Click = {
	
	
	Try
	{
		$OutputFile = Read-Host "Enter the location and name for the CSV. EX: C:\Scripts\UserPerms.csv"
		$textboxDetails.Text = "https://gallery.technet.microsoft.com/scriptcenter/Export-mailbox-permissions-d12a1d28"
		
		#Main
		Function Main
		{
			
			
			
			#Prepare Output file with headers
			Out-File -FilePath $OutputFile -InputObject "UserPrincipalName,ObjectWithAccess,ObjectType,AccessType,Inherited,AllowOrDeny" -Encoding UTF8
			
			
			$objUsers = get-mailbox -ResultSize Unlimited | Select-Object UserPrincipalName
			
			
			#Iterate through all users	
			Foreach ($objUser in $objUsers)
			{
				#Connect to the users mailbox
				$objUserMailbox = get-mailboxpermission -Identity $($objUser.UserPrincipalName) | Select-Object User, AccessRights, Deny, IsInherited
				
				#Prepare UserPrincipalName variable
				$strUserPrincipalName = $objUser.UserPrincipalName
				
				#Loop through each permission
				foreach ($objPermission in $objUserMailbox)
				{
					#Get the remaining permission details (We're only interested in real users, not built in system accounts/groups)
					if (($objPermission.user.tolower().contains("\domain admin")) -or ($objPermission.user.tolower().contains("\enterprise admin")) -or ($objPermission.user.tolower().contains("\organization management")) -or ($objPermission.user.tolower().contains("\administrator")) -or ($objPermission.user.tolower().contains("\exchange servers")) -or ($objPermission.user.tolower().contains("\public folder management")) -or ($objPermission.user.tolower().contains("nt authority")) -or ($objPermission.user.tolower().contains("\exchange trusted subsystem")) -or ($objPermission.user.tolower().contains("\discovery management")) -or ($objPermission.user.tolower().contains("s-1-5-21")))
					{ }
					Else
					{
						$objRecipient = (get-recipient $($objPermission.user) -EA SilentlyContinue)
						
						if ($objRecipient)
						{
							$strUserWithAccess = $($objRecipient.DisplayName) + " (" + $($objRecipient.PrimarySMTPAddress) + ")"
							$strObjectType = $objRecipient.RecipientType
						}
						else
						{
							$strUserWithAccess = $($objPermission.user)
							$strObjectType = "Other"
						}
						
						$strAccessType = $($objPermission.AccessRights) -replace ",", ";"
						
						if ($objPermission.Deny -eq $true)
						{
							$strAllowOrDeny = "Deny"
						}
						else
						{
							$strAllowOrDeny = "Allow"
						}
						
						$strInherited = $objPermission.IsInherited
						
						#Prepare the user details in CSV format for writing to file
						$strUserDetails = "$strUserPrincipalName,$strUserWithAccess,$strObjectType,$strAccessType,$strInherited,$strAllowOrDeny"
						
						$TextboxResults.Text = $strUserDetails
						
						#Append the data to file
						Out-File -FilePath $OutputFile -InputObject $strUserDetails -Encoding UTF8 -append
					}
				}
			}
			
			
		}
		
		# Start script
		. Main
		$TextboxResults.Text = ""
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$addFullAccessPermissionsToASharedMailboxToolStripMenuItem_Click = {
	$SharedMailboxFullAccess = Read-Host "Enter the name of the shared mailbox"
	$GrantFullAccesstoSharedMailbox = Read-Host "Enter the UPN of the user that will have full access"
	Try
	{
		$TextboxResults.Text = "Granting Full Access permissions to $GrantFullAccesstoSharedMailbox for the $SharedMailboxFullAccess shared mailbox..."
		$textboxDetails.Text = "Add-MailboxPermission $SharedMailboxFullAccess -User $GrantFullAccesstoSharedMailbox -AccessRights FullAccess -InheritanceType All | Format-List"
		$TextboxResults.Text = Add-MailboxPermission $SharedMailboxFullAccess -User $GrantFullAccesstoSharedMailbox -AccessRights FullAccess -InheritanceType All | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getSharedMailboxPermissionsToolStripMenuItem_Click = {
	$SharedMailboxPermissionsList = Read-Host "Enter the name of the Shared Mailbox"
	Try
	{
		$TextboxResults.Text = "Getting Send As permissions for $SharedMailboxPermissionsList..."
		$textboxDetails.Text = "Get-RecipientPermission $SharedMailboxPermissionsList | Where-Object { (`$_.Trustee -notlike 'NT AUTHORITY\SELF') } | Sort-Object Trustee | Format-Table Trustee, AccessControlType, AccessRights -AutoSize"
		#$TextboxResults.Text = Get-RecipientPermission $SharedMailboxPermissionsList | Format-List | Out-String
		$TextboxResults.Text = Get-RecipientPermission $SharedMailboxPermissionsList | Where-Object { ($_.Trustee -notlike "NT AUTHORITY\SELF") } | Sort-Object Trustee | Format-Table Trustee, AccessControlType, AccessRights -AutoSize| Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getSharedMailboxFullAccessPermissionsToolStripMenuItem_Click = {
	$SharedMailboxFullAccessPermissionsList = Read-Host "Enter the name of the Shared Mailbox"
	Try
	{
		$TextboxResults.Text = "Getting Full Access permissions for $SharedMailboxFullAccessPermissionsList..."
		$textboxDetails.Text = "Get-MailboxPermission $SharedMailboxFullAccessPermissionsList | Where-Object { (`$_.User -notlike 'NT AUTHORITY\SELF'') } | Sort-Object Identity | Format-Table Identity, User, AccessRights -AutoSize"
		$TextboxResults.Text = Get-MailboxPermission $SharedMailboxFullAccessPermissionsList | Where-Object { ($_.User -notlike "NT AUTHORITY\SELF") } | Sort-Object Identity | Format-Table Identity, User, AccessRights -AutoSize | Out-String

	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$addSendAsAccessToASharedMailboxToolStripMenuItem_Click = {
	$SharedMailboxSendAsAccess = Read-Host "Enter the name of the shared mailbox"
	$SharedMailboxSendAsUser = Read-Host "Enter the UPN of the user"
	Try
	{
		$TextboxResults.Text = "Getting Send As permissions for $SharedMailboxSendAsAccess..."
		$textboxDetails.Text = "Add-RecipientPermission $SharedMailboxSendAsAccess -Trustee $SharedMailboxSendAsUser -AccessRights SendAs | Format-List"
		$TextboxResults.Text = Add-RecipientPermission $SharedMailboxSendAsAccess -Trustee $SharedMailboxSendAsUser -AccessRights SendAs | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}



###CONTACTS###

$createANewMailContactToolStripMenuItem_Click = {
	$ContactFirstName = Read-Host "Enter the contacts first name"
	$ContactsLastName = Read-Host "Enter the contacts last name"
	$ContactName = $ContactFirstName + " " + $ContactsLastName
	$ContactExternalEmail = Read-Host "Enter external email for $ContactName"
	Try
	{
		$TextboxResults.Text = "Creating a new contact $ContactName"
		$textboxDetails.Text = "New-MailContact -Name $ContactName -FirstName $ContactFirstName -LastName $ContactsLastName -ExternalEmailAddress $ContactExternalEmail"
		New-MailContact -Name $ContactName -FirstName $ContactFirstName -LastName $ContactsLastName -ExternalEmailAddress $ContactExternalEmail
		$TextboxResults.Text = Get-MailContact -Identity $ContactName | Format-List DisplayName, EmailAddresses, PrimarySmtpAddress, ExternalEmailAddress, RecipientType | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getAllContactsToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all contacts..."
		$textboxDetails.Text = "Get-MailContact | Sort-Object DisplayName | Format-Table DisplayName, EmailAddresses, PrimarySmtpAddress, ExternalEmailAddress, RecipientType -AutoSize"
		$TextboxResults.Text = Get-MailContact | Sort-Object DisplayName | Format-Table DisplayName, EmailAddresses, PrimarySmtpAddress, ExternalEmailAddress, RecipientType -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getDetailedInfoForAContactToolStripMenuItem_Click = {
	$DetailedInfoForContact = Read-Host "Enter the contact name, displayname, alias or email address"
	Try
	{
		$TextboxResults.Text = "Getting detailed info for $DetailedInfoForContact..."
		$textboxDetails.Text = "Get-MailContact -Identity $DetailedInfoForContact | Format-List"
		$TextboxResults.Text = Get-MailContact -Identity $DetailedInfoForContact | Format-List | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$removeAContactToolStripMenuItem_Click = {
	$RemoveMailContact = Read-Host "Enter the contact name, displayname, alias or email address"
	Try
	{
		$TextboxResults.Text = "Removing contact $RemoveMailContact..."
		$textboxDetails.Text = "Remove-MailContact -Identity $RemoveMailContact"
		Remove-MailContact -Identity $RemoveMailContact
		$TextboxResults.Text = "Getting list of all contacts..."
		$TextboxResults.Text = Get-MailContact | Sort-Object DisplayName | Format-Table DisplayName, EmailAddresses, PrimarySmtpAddress, ExternalEmailAddress, RecipientType -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

	#Global Address List

$hideContactFromGALToolStripMenuItem_Click = {
	$HideGALMailContact = Read-Host "Enter the contact name, displayname, alias or email address"
	Try
	{
		$TextboxResults.Text = "Hiding $HideGALMailContact from the GAL..."
		$textboxDetails.Text = "Set-MailContact -Identity $HideGALMailContact -HiddenFromAddressListsEnabled `$True"
		Set-MailContact -Identity $HideGALMailContact -HiddenFromAddressListsEnabled $True
		$TextboxResults.Text = Get-MailContact -Identity $HideGALMailContact | Format-List DisplayName, HiddenFromAddressListsEnabled | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$unhideContactFromGALToolStripMenuItem_Click = {
	$unHideGALMailContact = Read-Host "Enter the contact name, displayname, alias or email address"
	Try
	{
		$TextboxResults.Text = "unhiding $unHideGALMailContact from the GAL..."
		$textboxDetails.Text = "Set-MailContact -Identity $unHideGALMailContact -HiddenFromAddressListsEnabled `$False"
		Set-MailContact -Identity $unHideGALMailContact -HiddenFromAddressListsEnabled $False
		$TextboxResults.Text = Get-MailContact -Identity $unHideGALMailContact | Format-List DisplayName, HiddenFromAddressListsEnabled | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getGALStatusForAllUsersToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting GAL status for all users..."
		$textboxDetails.Text = "Get-MailContact | Sort-Object DisplayName | Format-Table DisplayName, HiddenFromAddressListsEnabled -AutoSize"
		$TextboxResults.Text = Get-MailContact | Sort-Object DisplayName | Format-Table DisplayName, HiddenFromAddressListsEnabled -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getContactsHiddenFromGALToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all users that are hidden from the GAL..."
		$textboxDetails.Text = "Get-MailContact | Where-Object { `$_.HiddenFromAddressListsEnabled -like 'True' } | Sort-Object DisplayName | Format-Table DisplayName, HiddenFromAddressListsEnabled -AutoSize"
		$TextboxResults.Text = Get-MailContact | Where-Object { $_.HiddenFromAddressListsEnabled -like "True" } | Sort-Object DisplayName | Format-Table DisplayName, HiddenFromAddressListsEnabled -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getContactsNotHiddenFromGALToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all users that not are hidden from the GAL"
		$textboxDetails.Text = "Get-MailContact | Where-Object { `$_.HiddenFromAddressListsEnabled -like 'False' } | Sort-Object DisplayName | Format-Table DisplayName, HiddenFromAddressListsEnabled -AutoSize"
		$TextboxResults.Text = Get-MailContact | Where-Object { $_.HiddenFromAddressListsEnabled -like "False" } | Sort-Object DisplayName | Format-Table DisplayName, HiddenFromAddressListsEnabled -AutoSize | Out-String
	}
	Catch
	{
		$TextboxResults.Text = ""
		$textboxDetails.Text = ""
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}



###FILE###

	#About

$aboutToolStripMenuItem_Click = {
	$TextboxResults.Text = "                 o365 Administration Center v2.0.3
	
HOW TO USE

To start, click the Connect to Office 365 button. This will connect you to Exchange Online using Remote PowerShell. 
Once you are connected the button will grey out and the form title will change to -CONNECTED TO O365-

The TextBox will display all output for each command. 
If nothing appears and there was no error then the result was null. 
The Textbox also serves as input, passing your own commands to PowerShell with the result populating in the same Textbox. 
To run your own command simply clear the Textbox and enter in your command and press the Run Command button or press Enter on your keyboard.

You can also export the results to a file using the Export to File button. 
The Textbox also allows copy and paste. 
Closing the form properly removes all PSSessions"
	
}

	#Pre-Reqs

$prerequisitesToolStripMenuItem_Click = {
	$TextboxResults.Text = "
Windows Versions:
Windows 7 Service Pack 1 (SP1) or higher
Windows Server 2008 R2 SP1 or higher

You need to install the Microsoft.NET Framework 4.5 or later and then Windows Management Framework 3.0 or later. 

MS-Online module needs to be installed. Install the MSOnline Services Sign In Assistant: 
https://www.microsoft.com/en-us/download/details.aspx?id=41950 

Azure Active Directory Module for Windows PowerShell needs to be installed: http://go.microsoft.com/fwlink/p/?linkid=236297

The Microsoft Online Services Sign-In Assistant provides end user sign-in capabilities to Microsoft Online Services, such as Office 365.

Windows PowerShell needs to be configured to run scripts, and by default, it isn't. 
To enable Windows PowerShell to run scripts, run the following command in an elevated Windows PowerShell window (a Windows PowerShell window you open by selecting Run as administrator): 
Set-ExecutionPolicy Unrestricted

PowerShell v3 or higher"
	
	
	
}

	#Buy Me A Beer
$buyMeABeerToolStripMenuItem_Click = {
	$textboxDetails.Text = ""
	$TextboxResults.Text = ""
	Start-Process -FilePath https://www.paypal.me/bwya77
	#$TextboxResults.Text = "https://www.paypal.me/bwya77


#Thank You!"
}





###JUNK ITEMS###

$TextboxResults_TextChanged={
	
}

$menustrip1_ItemClicked=[System.Windows.Forms.ToolStripItemClickedEventHandler]{
	#Event Argument: $_ = [System.Windows.Forms.ToolStripItemClickedEventArgs]
}

$WhitelistToolStripMenuItem_Click={
	
}

$organizationCustomizationToolStripMenuItem_Click={
	
}

$getMailMalwareReportToolStripMenuItem_Click = {
	
}

$securityGroupsToolStripMenuItem_Click={
	
}

