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
	$ButtonExit.Text = "Exit"
	
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
	$ButtonDisconnectFromPartner.Enabled = $False
	
	#Alphabitcally sorts combobox
	$PartnerComboBox.Sorted = $True
	
	#Disables word wrap on the text box
	$TextboxResults.WordWrap = $false
	
	#Enables vertical and horizontal scrollbars
	$TextboxResults.ScrollBars = 'Both'
	
}

	#Buttons

$ButtonExit_Click= {
		#Disconnects O365 Session
		Get-PSSession | Remove-PSSession
		
		<# Creates a pop up box telling the user they are disconnected from the o365 session. This is commented out as it will show True every time as the command will never error out even if there 
		is no session to disconnect from #>
		#[void][System.Windows.Forms.MessageBox]::Show("You are disconnected from O365", "Message")
	}

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
	$PartnerConnectButton.Text = "Connect to O365 Partner"
	$FormO365AdministrationCenter.Text = "O365 Administration Center"
	$PartnerComboBox.Enabled = $True
	$TextboxResults.Text = ""
	
}

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
		
	}
	Catch
	{
		$TextboxResults.Text = ""
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
		#$PartnerConnectButton.Text = "Connected to O365 Partner"
		#Sets custom form text
		$FormO365AdministrationCenter.Text = "-Connected to O365 Partner: "+ $PartnerComboBox.SelectedItem.Name+"-"
		$TextboxResults.Text = ""
		$PartnerComboBox.Enabled = $false
		$ButtonDisconnectFromPartner.Enabled = $true
	}
	catch
	{
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
			[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		} 
}

$ButtonRunCustomCommand_Click = {
	$userinput = $TextboxResults.text
			try
			{
				#Takes the user input to a variable and passes it to the shell
				$TextboxResults.Text = "Running command $userinput..."
				$TextboxResults.text = Invoke-Expression $userinput | Format-List | Out-String
			}
			Catch
			{
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
		Set-MailboxAutoReplyConfiguration -Identity $OOOautoreplyUser -AutoReplyState Enabled -ExternalMessage $OOOExternal -InternalMessage $OOOInternal
		$TextboxResults.Text = Get-MailboxAutoReplyConfiguration -Identity $OOOautoreplyUser | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getListOfUsersToolStripMenuItem_Click = {
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting list of users..."
		$TextboxResults.text = Get-MSOLUser | Select-Object DisplayName, UserPrincipalName | Format-Table -AutoSize | Out-String
		}
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue )
		{
		$TextboxResults.Text = "Getting list of users..."
		$TextboxResults.text = Get-MSOLUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Select-Object DisplayName, UserPrincipalName  | Format-Table -AutoSize | Out-String
		}
	Else
		{
		[System.Windows.Forms.MessageBox]::Show("Could not get a list of users", "Error")
		$TextboxResults.Text = ""
		}
}

$getDetailedInfoForAUserToolStripMenuItem_Click = {
	$DetailedInfoUser = Read-Host "Enter the UPN of the user you want more information about"
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Getting detailed info for $DetailedInfoUser..."
			$TextboxResults.text = Get-MsolUser -UserPrincipalName $DetailedInfoUser | Format-List | Out-String
		}
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Getting detailed info for $DetailedInfoUser..."
			$TextboxResults.text = Get-MsolUser -UserPrincipalName $DetailedInfoUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List | Out-String
		}
		Else
		{
		[System.Windows.Forms.MessageBox]::Show("Could not get detailed info for $DetailedInfoUser", "Error")
		$TextboxResults.Text = ""
		}
}

$changeUsersLoginNameToolStripMenuItem_Click = {
	$UserChangeUPN = Read-Host "What user would you like to change their login name for? Enter their UPN"
	$NewUserUPN = Read-Host "What would you like the new username to be?"
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Changing $UserChangeUPN UPN to $NewUserUPN..."
			Set-MsolUserPrincipalname -UserPrincipalName $UserChangeUPN -NewUserPrincipalName $NewUserUPN
			$TextboxResults.text = Get-MSOLUser -UserPrincipalName $NewUserUPN | Format-List UserPrincipalName | Out-String
		}
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Changing $UserChangeUPN UPN to $NewUserUPN..."
			Set-MsolUserPrincipalname -UserPrincipalName $UserChangeUPN -TenantId $PartnerComboBox.SelectedItem.TenantID -NewUserPrincipalName $NewUserUPN
			$TextboxResults.text = Get-MSOLUser -UserPrincipalName $NewUserUPN -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName | Out-String
		}
		Else
		{
		[System.Windows.Forms.MessageBox]::Show("Could not change the login name for $UserChangeUPN", "Error")
		$TextboxResults.Text = ""
		}
}

$deleteAUserToolStripMenuItem_Click = {
	$DeleteUser = Read-Host "Enter the UPN of the user you want to delete"
		#What to do if connected to main o365 account
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Deleting $DeleteUser..."
			$TextboxResults.text = Remove-MsolUser –UserPrincipalName $DeleteUser | Format-List UserPrincipalName | Out-String
		}
		#What to do if connected to partner account
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Deleting $DeleteUser..."
			$TextboxResults.text = Remove-MsolUser –UserPrincipalName $DeleteUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName | Out-String
		}
		Else
		{
		[System.Windows.Forms.MessageBox]::Show("Could not delete $DeleteUser", "Error")
		$TextboxResults.Text = ""
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
			$TextboxResults.text = New-MsolUser -UserPrincipalName $NewUser -FirstName $Firstname -LastName $LastName -DisplayName $DisplayName | Format-List | Out-String
		}
		#What to do if connected to partner account
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Creating user $NewUser..."
			$TextboxResults.text = New-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $NewUser -FirstName $Firstname -LastName $LastName -DisplayName $DisplayName | Format-List | Out-String
		}
		Else
		{
		[System.Windows.Forms.MessageBox]::Show("Could not create the new user $newuser", "Error")
		$TextboxResults.Text = ""
		}
}

$disableUserAccountToolStripMenuItem_Click = {
	$BlockUser = Read-Host "Enter the UPN of the user you want to disable"
		#What to do if connected to main o365 account
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Disabling $BlockUser..."
			Set-MsolUser -UserPrincipalName $BlockUser -blockcredential $True 
			$TextboxResults.text = Get-MsolUser -UserPrincipalName $BlockUser | Format-List DisplayName, BlockCredential | Out-String
		}
		#What to do if connected to partner account
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Disabling $BlockUser..."
			Set-MsolUser -UserPrincipalName $BlockUser -blockcredential $True -TenantId $PartnerComboBox.SelectedItem.TenantID
			$TextboxResults.text = Get-MsolUser -UserPrincipalName $BlockUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List DisplayName, BlockCredential | Out-String
		}
		Else
		{
		[System.Windows.Forms.MessageBox]::Show("Could not disable $BlockUser", "Error")
		$TextboxResults.Text = ""
		}
}

$enableAccountToolStripMenuItem_Click = {
	$EnableUser = Read-Host "Enter the UPN of the user you want to enable"
		#What to do if connected to main o365 account
		If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Enabling $EnableUser..."
			Set-MsolUser -UserPrincipalName $EnableUser -blockcredential $False
			$TextboxResults.text = Get-MsolUser -UserPrincipalName $EnableUser | Format-List DisplayName, BlockCredential | Out-String
		}
		#What to do if connected to partner account
		ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
		{
			$TextboxResults.Text = "Enabling $EnableUser..."
			Set-MsolUser -UserPrincipalName $EnableUser -blockcredential $False -TenantId $PartnerComboBox.SelectedItem.TenantID
			$TextboxResults.text = Get-MsolUser -UserPrincipalName $EnableUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List DisplayName, BlockCredential | Out-String
		}
		Else
		{
		[System.Windows.Forms.MessageBox]::Show("Could enable $EnableUser", "Error")
		$TextboxResults.Text = ""
		}
}

	#Quota

$getUserQuotaToolStripMenuItem_Click={
	$QuotaUser = Read-Host "Enter the Email of the user you want to view Quota information for"
	try
	{
		$TextboxResults.Text = "Getting user quota for $QuotaUser..."
		$TextboxResults.text = Get-Mailbox $QuotaUser | Format-List DisplayName, UserPrincipalName, *Quota | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getAllUsersQuotaToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting quota for all users..."
		$TextboxResults.text = Get-Mailbox | Format-List DisplayName, UserPrincipalName, *Quota | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
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
		Set-Mailbox $MailboxSetQuota -ProhibitSendReceiveQuota $ProhibitSendReceiveQuota -ProhibitSendQuota $ProhibitSendQuota -IssueWarningQuota $IssueWarningQuota
		$TextboxResults.text = Get-Mailbox $MailboxSetQuota | Format-List DisplayName, UserPrincipalName, *Quota | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$setMailboxQuotaForAllToolStripMenuItem_Click = {
	$ProhibitSendReceiveQuota2 = Read-Host "Enter (GB) the ProhibitSendReceiveQuota value (EX: 50GB) Max:50GB"
	$ProhibitSendQuota2 = Read-Host "Enter (GB) the ProhibitSendQuota value (EX: 48GB) Max:50GB"
	$IssueWarningQuota2 = Read-Host "Enter (GB) theIssueWarningQuota value (EX: 45GB) Max:50GB"
	Try
	{
		$TextboxResults.Text = "Setting quota for all... "
		Get-Mailbox | Set-Mailbox -ProhibitSendReceiveQuota $ProhibitSendReceiveQuota2 -ProhibitSendQuota $ProhibitSendQuota2 -IssueWarningQuota $IssueWarningQuota2
		$TextboxResults.text = Get-Mailbox | Format-List DisplayName, UserPrincipalName, *Quota | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

	#Licenses

$getLicensedUsersToolStripMenuItem_Click={
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all users with a license..."
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.isLicensed -eq "TRUE" } | Format-List DisplayName, Licenses | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all users with a license..."
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.isLicensed -eq "TRUE" } | Format-List DisplayName, Licenses | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("Could not get license information", "Error")
		$TextboxResults.Text = ""
	}
}

$displayAllUsersWithoutALicenseToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all users without a license..."
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.isLicensed -like "False" } | Format-List UserPrincipalName | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all users without a license..."
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.isLicensed -like "False" } | Format-List UserPrincipalName | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("Could not display users without a license", "Error")
		$TextboxResults.Text = ""
	}
	
}

$removeAllUnlicensedUsersToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Removing all users without a license..."
		Get-MsolUser | Where-Object { $_.isLicensed -ne "true" } | Remove-MsolUser -Force -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.isLicensed -like "False" } | Format-List UserPrincipalName | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Removing all users without a license..."
		Get-MsolUser -all -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.isLicensed -ne "true" } | Remove-MsolUser -Force -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.isLicensed -like "False" } | Format-List UserPrincipalName | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("Could not remove all unlicensed users", "Error")
		$TextboxResults.Text = ""
	}
}

$displayAllLicenseInfoToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all license information..."
		$TextboxResults.text = Get-MsolAccountSku | Select-Object -Property AccountSkuId, ActiveUnits, WarningUnits, ConsumedUnits, @{
			Name = 'Unused'
			Expression = {
				$_.ActiveUnits - $_.ConsumedUnits
			}
		} | Format-Table -AutoSize | Out-String
	}
    #What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all license information..."
		$TextboxResults.text = Get-MsolAccountSku -TenantId $PartnerComboBox.SelectedItem.TenantID | Select-Object -Property AccountSkuId, ActiveUnits, WarningUnits, ConsumedUnits, @{
			Name = 'Unused'
			Expression = {
				$_.ActiveUnits - $_.ConsumedUnits
			}
		} | Format-Table -AutoSize | Out-String
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
		$TextboxResults.text = Get-MsolAccountSku | Format-Table -AutoSize | Out-String
		$LicenseType = Read-Host "Enter the AccountSku of the License you want to assign to this user"
		$TextboxResults.Text = "Adding $LicenseType license to $LicenseUserAdd..."
		Set-MsolUser -UserPrincipalName $LicenseUserAdd –UsageLocation $LicenseUserAddLocation
		Set-MsolUserLicense -UserPrincipalName $LicenseUserAdd -AddLicenses $LicenseType
		$TextboxResults.Text = Get-MsolUser -UserPrincipalName $LicenseUserAdd | Format-List DisplayName, Licenses | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$LicenseUserAdd = Read-Host "Enter the User Principal Name of the User you want to license"
		$LicenseUserAddLocation = Read-Host "Enter the 2 digit location code for the user. Example: US"
		$TextboxResults.text = Get-MsolAccountSku -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-Table -AutoSize | Out-String
		$LicenseType = Read-Host "Enter the AccountSku of the License you want to assign to this user"
		$TextboxResults.Text = "Adding $LicenseType license to $LicenseUserAdd..."
		Set-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $LicenseUserAdd –UsageLocation $LicenseUserAddLocation
		Set-MsolUserLicense -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $LicenseUserAdd -AddLicenses $LicenseType
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $LicenseUserAdd | Format-List DisplayName, Licenses | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
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
		Set-MsolUserLicense -UserPrincipalName $RemoveLicenseFromUser -RemoveLicenses $RemoveLicenseType
		$TextboxResults.Text = Get-MsolUser -UserPrincipalName $RemoveLicenseFromUser | Format-List DisplayName, Licenses | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$RemoveLicenseFromUser = Read-Host "Enter the User Principal Name of the user you want to remove a license from"
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $RemoveLicenseFromUser | Format-List DisplayName, Licenses | Out-String
		$RemoveLicenseType = Read-Host "Enter the AccountSku of the license you want to remove"
		$TextboxResults.Text = "Removing the $RemoveLicenseType license from $RemoveLicenseFromUser..."
		Set-MsolUserLicense -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $RemoveLicenseFromUser -RemoveLicenses $RemoveLicenseType
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $RemoveLicenseFromUser | Format-List DisplayName, Licenses | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
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
		Remove-MailboxFolderPermission -identity ${Calendaruser}:\calendar -user $Calendaruser2
		Add-MailboxFolderPermission -Identity ${Calendaruser}:\calendar -user $Calendaruser2 -AccessRights $level
		$TextboxResults.Text = Get-MailboxFolderPermission -Identity ${Calendaruser}:\calendar | Format-List Identity, FolderName, User, AccessRights, IsValid, ObjectSpace | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$viewUsersCalendarPermissionsToolStripMenuItem_Click = {
	$CalUserPermissions = Read-Host "What user would you like calendar permissions for?"
	Try
	{
		$TextboxResults.Text = "Getting $CalUserPermissions calendar permissions..."
		$TextboxResults.Text = Get-MailboxFolderPermission -Identity ${CalUserPermissions}:\calendar | Format-List Identity, FolderName, User, AccessRights, IsValid, ObjectSpace | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
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
		$users = Get-Mailbox | Select-Object -ExpandProperty Alias
		Foreach ($user in $users) { Add-MailboxFolderPermission ${user}:\Calendar -user $MasterUser -accessrights $level2 }﻿
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
$TextboxResults.Text = ""
}

$removeAUserFromAllCalendarsToolStripMenuItem_Click = {
	$RemoveUserFromAll = Read-Host "Enter the UPN of the user you want to remove from all calendars"
	try
	{
		$TextboxResults.Text = "Removing $RemoveUserFromAll from all users calendar..."
		$users = Get-Mailbox | Select-Object -ExpandProperty Alias
		Foreach ($user in $users) { Remove-MailboxFolderPermission ${user}:\Calendar -user $RemoveUserFromAll -Confirm:$false}﻿
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$removeAUserFromSomesonsCalendarToolStripMenuItem_Click = {
	$Calendaruserremove = Read-Host "Enter the UPN of the user whose calendar you want to remove access to"
	$Calendaruser2remove = Read-Host "Enter the UPN of the user who you want to remove access"
	try
	{
		$TextboxResults.Text = "Removing $Calendaruser2remove from $Calendaruserremove calendar..."
		Remove-MailboxFolderPermission -Identity ${Calendaruserremove}:\calendar -user $Calendaruser2remove
		$TextboxResults.Text = Get-MailboxFolderPermission -Identity ${Calendaruserremove}:\calendar | Format-List Identity, FolderName, User, AccessRights, IsValid, ObjectSpace | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Clutter

$disableClutterForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Disabling Clutter for all users..."
		$TextboxResults.text = Get-Mailbox | Set-Clutter -Enable $false | Format-List MailboxIdentity, IsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$enableClutterForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Enabling Clutter for all users..."
		$TextboxResults.text = Get-Mailbox | Set-Clutter -Enable $True | Format-List MailboxIdentity, IsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$enableClutterForAUserToolStripMenuItem_Click = {
	$UserEnableClutter = Read-Host "Which user would you like to enable Clutter for?"
	try
	{
		$TextboxResults.Text = "Enabling Clutter for $UserEnableClutter..."
		$TextboxResults.text = Get-Mailbox $UserEnableClutter | Set-Clutter -Enable $True | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$disableClutterForAUserToolStripMenuItem_Click = {
	$UserDisableClutter = Read-Host "Which user would you like to disable Clutter for?"
	try
	{
		$TextboxResults.Text = "Disabling Clutter for $UserDisableClutter..."
		$TextboxResults.text = Get-Mailbox $UserDisableClutter | Set-Clutter -Enable $False | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getClutterInfoForAUserToolStripMenuItem_Click = {
	$GetCluterInfoUser = Read-Host "What user would you like to view Clutter information about?"
	Try
	{
		$TextboxResults.Text = "Getting Clutter information for $GetCluterInfoUser..."
		$TextboxResults.Text = Get-Clutter -Identity $GetCluterInfoUser | Format-List MailboxIdentity, IsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Recycle Bin

$displayAllDeletedUsersToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all deleted users..."
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all deleted users..."
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("Could not display users without a license", "Error")
		$TextboxResults.Text = ""
	}
}

$deleteAllUsersInRecycleBinToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Deleting all users in the recycle bin..."
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin –Force | Format-List | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Deleting all users in the recycle bin..."
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin –Force | Format-List | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("Could not display users without a license", "Error")
		$TextboxResults.Text = ""
	}
	
}

$deleteSpecificUsersInRecycleBinToolStripMenuItem_Click = {
	$DeletedUserRecycleBin = Read-Host "Please enter the User Principal Name of the user you want to permanently delete"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Deleting  $DeletedUserRecycleBin from the recycle bin..."
		Remove-MsolUser -UserPrincipalName $DeletedUserRecycleBin -RemoveFromRecycleBin -Force
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Deleting  $DeletedUserRecycleBin from the recycle bin..."
		Remove-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $DeletedUserRecycleBin -RemoveFromRecycleBin -Force
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("Could not display users without a license", "Error")
		$TextboxResults.Text = ""
	}
}

$restoreDeletedUserToolStripMenuItem_Click = {
	$RestoredUserFromRecycleBin = Read-Host "Enter the User Principal Name of the user you want to restore"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Restoring $RestoredUserFromRecycleBin from the recycle bin..."
		Restore-MsolUser –UserPrincipalName $RestoredUserFromRecycleBin -AutoReconcileProxyConflicts
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Restoring $RestoredUserFromRecycleBin from the recycle bin..."
		Restore-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID –UserPrincipalName $RestoredUserFromRecycleBin -AutoReconcileProxyConflicts
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$restoreAllDeletedUsersToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Restoring all deleted users..."
		Get-MsolUser -ReturnDeletedUsers | Restore-MsolUser
		$TextboxResults.Text = "Users that were deleted have now been restored"
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Restoring all deleted users..."
		Get-MsolUser -ReturnDeletedUsers -TenantId $PartnerComboBox.SelectedItem.TenantID | Restore-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.Text = "Users that were deleted have now been restored"
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Quarentine

$getQuarantineBetweenDatesToolStripMenuItem_Click = {
	$StartDateQuarentine = Read-Host "Enter the beginning date. (Format MM/DD/YYYY)"
	$EndDateQuarentine = Read-Host "Enter the end date. (Format MM/DD/YYYY)"
	try
	{
		$TextboxResults.Text = "Getting quarantine between $StartDateQuarentine and $EndDateQuarentine..."
		$TextboxResults.Text = Get-QuarantineMessage -StartReceivedDate $StartDateQuarentine -EndReceivedDate $EndDateQuarentine | Format-List ReceivedTime, SenderAddress, RecipientAddress, Subject, Size, Type, Expires, QuarantinedUser, ReleasedUser, Direction | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getQuarantineFromASpecificUserToolStripMenuItem_Click = {
	$QuarentineFromUser = Read-Host "Enter the email address you want to see quarentine from"
	try
	{
		$TextboxResults.Text = "Getting quarantine sent from $QuarentineFromUser ..."
		$TextboxResults.Text = Get-QuarantineMessage -SenderAddress $QuarentineFromUser | Format-List ReceivedTime, SenderAddress, RecipientAddress, Subject, Size, Type, Expires, QuarantinedUser, ReleasedUser, Direction | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getQuarantineToASpecificUserToolStripMenuItem_Click = {
	$QuarentineInfoForUser = Read-Host "Enter the email of the user you want to view quarantine for"
	try
	{
		$TextboxResults.Text = "Getting quarantine sent to $QuarentineInfoForUser..."
		$TextboxResults.Text = Get-QuarantineMessage -RecipientAddress $QuarentineInfoForUser | Format-List ReceivedTime, SenderAddress, Subject, Size, Type, Expires, QuarantinedUser, ReleasedUser, Direction | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Passwords

$enableStrongPasswordForAUserToolStripMenuItem_Click = {
	$UserEnableStrongPasswords = Read-Host "Enter the User Principal Name of the user you want to enable strong password policy for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Enabling strong password policy for $UserEnableStrongPasswords..."
		Set-MsolUser -UserPrincipalName $UserEnableStrongPasswords -StrongPasswordRequired $True
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $UserEnableStrongPasswords | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
		
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Enabling strong password policy for $UserEnableStrongPasswords..."
		Set-MsolUser -UserPrincipalName $UserEnableStrongPasswords -StrongPasswordRequired $True -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $UserEnableStrongPasswords -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
		
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getAllUsersStrongPasswordPolicyInfoToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting strong password policy for all users..."
		$TextboxResults.text = Get-MsolUser | Format-List userprincipalname, strongpasswordrequired | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting strong password policy for all users..."
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List userprincipalname, strongpasswordrequired | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$disableStrongPasswordsForAUserToolStripMenuItem_Click = {
	$UserdisableStrongPasswords = Read-Host "Enter the User Principal Name of the user you want to disable strong password policy for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Disabling strong password policy for $UserdisableStrongPasswords..."
		Set-MsolUser -UserPrincipalName $UserdisableStrongPasswords -StrongPasswordRequired $False
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $UserdisableStrongPasswords | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
		
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Disabling strong password policy for $UserdisableStrongPasswords..."
		Set-MsolUser -UserPrincipalName $UserdisableStrongPasswords -StrongPasswordRequired $False -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $UserdisableStrongPasswords -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$enableStrongPasswordsForAllToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Enabling strong password policy for all users..."
		Get-MsolUser | Set-MsolUser -StrongPasswordRequired $True
		$TextboxResults.text = Get-MsolUser | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Enabling strong password policy for all users..."
		Get-MsolUser | Set-MsolUser -StrongPasswordRequired $True -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$disableStrongPasswordsForAllToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Disabling strong password policy for all users..."
		Get-MsolUser | Set-MsolUser -StrongPasswordRequired $False
		$TextboxResults.text = Get-MsolUser | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Disabling strong password policy for all users..."
		Get-MsolUser | Set-MsolUser -StrongPasswordRequired $False -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$resetPasswordForAUserToolStripMenuItem1_Click = {
	$ResetPasswordUser = Read-Host "Who user would you like to reset the password for?"
	$NewPassword = Read-Host "What would you like the new password to be?"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Resetting $ResetPasswordUser password to $NewPassword..."
		Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser –NewPassword $NewPassword -ForceChangePassword $False
		$TextboxResults.Text = "The password for $ResetPasswordUser has been set to $NewPassword"
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Resetting $ResetPasswordUser password to $NewPassword..."
		Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser –NewPassword $NewPassword -ForceChangePassword $False -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.Text = "The password for $ResetPasswordUser has been set to $NewPassword"
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$setPasswordToNeverExpireForAllToolStripMenuItem1_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to never expire for all..."
		Get-MsolUser | Set-MsolUser –PasswordNeverExpires $True
		$TextboxResults.text = Get-MSOLUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to never expire for all..."
		Get-MsolUser | Set-MsolUser –PasswordNeverExpires $True -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MSOLUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$setPasswordToExpireForAllToolStripMenuItem1_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to expire for all..."
		Get-MsolUser | Set-MsolUser –PasswordNeverExpires $False
		$TextboxResults.text = Get-MSOLUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to expire for all..."
		Get-MsolUser | Set-MsolUser –PasswordNeverExpires $False -TenantId $PartnerComboBox.SelectedItem.TenantID
		$TextboxResults.text = Get-MSOLUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$resetPasswordForAllToolStripMenuItem_Click = {
	$SetPasswordforAll = Read-Host "What password would you like to set for all users?"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Resetting all users passwords to $SetPasswordforAll..."
		Get-MsolUser | ForEach-Object{ Set-MsolUserPassword -userPrincipalName $_.UserPrincipalName –NewPassword $SetPasswordforAll -ForceChangePassword $False }
		$TextboxResults.Text = "Password for all users has been set to $SetPasswordforAll"
		
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Resetting all users passwords to $SetPasswordforAll..."
		Get-MsolUser | ForEach-Object{ Set-MsolUserPassword -TenantId $PartnerComboBox.SelectedItem.TenantID -userPrincipalName $_.UserPrincipalName –NewPassword $SetPasswordforAll -ForceChangePassword $False }
		$TextboxResults.Text = "Password for all users has been set to $SetPasswordforAll"
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$setATemporaryPasswordForAllToolStripMenuItem_Click = {
	$SetTempPasswordforAll = Read-Host "What password would you like to set for all users?"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting $SetTempPasswordforAll as the temporary password for all users..."
		Get-MsolUser | Set-MsolUserPassword –NewPassword $SetTempPasswordforAll -ForceChangePassword $True
		$TextboxResults.Text = "Temporary password has been set to $SetTempPasswordforAll Please note that users will be prompted to change it upon logon"
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting $SetTempPasswordforAll as the temporary password for all users..."
		Get-MsolUser | Set-MsolUserPassword -TenantId $PartnerComboBox.SelectedItem.TenantID –NewPassword $SetTempPasswordforAll -ForceChangePassword $True
		$TextboxResults.Text = "Temporary password has been set to $SetTempPasswordforAll Please note that users will be prompted to change it upon logon"
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$TemporaryPasswordForAUserToolStripMenuItem_Click = {
	$ResetPasswordUser2 = Read-Host "Who user would you like to reset the password for?"
	$NewPassword2 = Read-Host "What would you like the new password to be?"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting $NewPassword2 as the temporary password for $ResetPasswordUser2..."
		Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser2 –NewPassword $NewPassword2 -ForceChangePassword $True
		$TextboxResults.Text = "Temporary password has been set to $NewPassword2 Please note that $ResetPasswordUser2 will be prompted to change it upon logon"
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting $NewPassword2 as the temporary password for $ResetPasswordUser2..."
		Set-MsolUserPassword -TenantId $PartnerComboBox.SelectedItem.TenantID –UserPrincipalName $ResetPasswordUser2 –NewPassword $NewPassword2 -ForceChangePassword $True
		$TextboxResults.Text = "Temporary password has been set to $NewPassword2 Please note that $ResetPasswordUser2 will be prompted to change it upon logon"
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getPasswordResetDateForAUserToolStripMenuItem_Click = {
	$GetPasswordInfoUser = Read-Host "Enter the UPN of the user you want to view the password last changed date for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password reset date for $GetPasswordInfoUser..."
		$TextboxResults.Text = Get-MsolUser -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password reset date for $GetPasswordInfoUser..."
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getPasswordLastResetDateForAllToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password reset date for all users..."
		$TextboxResults.Text = Get-MsolUser | Format-List UserPrincipalName, lastpasswordchangetimestamp | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password reset date for all users..."
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List UserPrincipalName, lastpasswordchangetimestamp | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$setPasswordToExpireForAUserToolStripMenuItem_Click = {
	$PasswordtoExpireforUser = Read-Host "Enter the UPN of the user you want the password to expire for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to expire for $PasswordtoExpireforUser..."
		Set-MsolUser -UserPrincipalName $PasswordtoExpireforUser –PasswordNeverExpires $False
		$TextboxResults.text = Get-MSOLUser -UserPrincipalName $PasswordtoExpireforUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to expire for $PasswordtoExpireforUser..."
		Set-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $PasswordtoExpireforUser –PasswordNeverExpires $False
		$TextboxResults.text = Get-MSOLUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $PasswordtoExpireforUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$setPasswordToNeverExpireForAUserToolStripMenuItem_Click = {
	$PasswordtoNeverExpireforUser = Read-Host "Enter the UPN of the user you want the password to expire for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to never expire for $PasswordtoNeverExpireforUser..."
		Set-MsolUser -UserPrincipalName $PasswordtoNeverExpireforUser –PasswordNeverExpires $True
		$TextboxResults.text = Get-MSOLUser -UserPrincipalName $PasswordtoNeverExpireforUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Setting password to never expire for $PasswordtoNeverExpireforUser..."
		Set-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $PasswordtoNeverExpireforUser –PasswordNeverExpires $True
		$TextboxResults.text = Get-MSOLUser -TenantId $PartnerComboBox.SelectedItem.TenantID -UserPrincipalName $PasswordtoNeverExpireforUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getUsersWhosPasswordNeverExpiresToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting users where the password is set to never expire..."
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.PasswordNeverExpires -eq $True } | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting users where the password is set to never expire..."
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.PasswordNeverExpires -eq $True } | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getUsersWhosPasswordWillExpireToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting users where the password is set to expire..."
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.PasswordNeverExpires -eq $False } | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting users where the password is set to expire..."
		$TextboxResults.text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID | Where-Object { $_.PasswordNeverExpires -eq $False } | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getPasswordLastResetDateForAUserToolStripMenuItem_Click = {
	$GetPasswordInfoUser = Read-Host "Enter the UPN of the user you want to view the password last changed date for"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password reset date for $GetPasswordInfoUser..."
		$TextboxResults.Text = Get-MsolUser -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password reset date for $GetPasswordInfoUser..."
		$TextboxResults.Text = Get-MsolUser -TenantId $PartnerComboBox.SelectedItem.TenantID -userprincipalname $GetPasswordInfoUser | Format-List UserPrincipalName, lastpasswordchangetimestamp | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getUsersNextPasswordResetDateToolStripMenuItem_Click = {
	$NextUserResetDateUser = Read-Host "Enter the UPN of the user"
	$VarDate = Read-Host "Enter days before passwords expires. EX: 90"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = (get-msoluser -userprincipalname $NextUserResetDateUser).lastpasswordchangetimestamp.adddays($VarDate) | Format-List DateTime | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = (get-msoluser -TenantId $PartnerComboBox.SelectedItem.TenantID -userprincipalname $NextUserResetDateUser).lastpasswordchangetimestamp.adddays($VarDate) | Format-List DateTime | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

	#Mailbox Permissions

$addFullPermissionsToAMailboxToolStripMenuItem_Click = {
	$mailboxAccess = read-host "Mailbox you want to give full-access to"
	$mailboxUser = read-host "Enter the UPN of the user that will have full access"
	try
	{
		$TextboxResults.Text = "Assigning full access permissions to $mailboxUser for the account $mailboxAccess..."
		$TextboxResults.text = Add-MailboxPermission $mailboxAccess -User $mailboxUser -AccessRights FullAccess -InheritanceType All | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$addSendAsPermissionToAMailboxToolStripMenuItem_Click = {
	$SendAsAccess = read-host "Mailbox you want to give Send As access to"
	$mailboxUserAccess = read-host "Enter the UPN of the user that will have Send As access"
	try
	{
		$TextboxResults.Text = "Assigning Send-As access to $mailboxUserAccess for the account $SendAsAccess..."
		$TextboxResults.text = Add-RecipientPermission $SendAsAccess -Trustee $mailboxUserAccess -AccessRights SendAs -Confirm:$False | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$assignSendOnBehalfPermissionsForAMailboxToolStripMenuItem_Click = {
	$SendonBehalfof = read-host "Mailbox you want to give Send As access to"
	$mailboxUserSendonBehalfAccess = read-host "Enter the UPN of the user that will have Send As access"
	try
	{
		$TextboxResults.Text = "Assigning Send On Behalf of permissions to $mailboxUserSendonBehalfAccess for the account $SendonBehalfof..."
		Set-Mailbox -Identity $SendonBehalfof -GrantSendOnBehalfTo $mailboxUserSendonBehalfAccess
		$TextboxResults.text = Get-Mailbox -Identity $SendonBehalfof | Format-List Identity, GrantSendOnBehalfTo | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$displayMailboxPermissionsForAUserToolStripMenuItem_Click = {
	$MailboxUserFullAccessPermission = Read-Host "Enter the UPN of the user want to view Full Access permissions for"
	try
	{
		$TextboxResults.Text = "Getting Full Access permissions for $MailboxUserFullAccessPermission..."
		$TextboxResults.text = Get-MailboxPermission $MailboxUserFullAccessPermission | Where-Object { ($_.IsInherited -eq $False) -and -not ($_.User -like "NT AUTHORITY\SELF") } | Format-List AccessRights, Deny, InheritanceType, User, Identity, IsInherited, IsValid | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$displaySendAsPermissionForAMailboxToolStripMenuItem_Click = {
	$MailboxUserSendAsPermission = Read-Host "Enter the UPN of the user you want to view Send As permissions for"
	try
	{
		$TextboxResults.Text = "Getting Send As Permissions for $MailboxUserSendAsPermission..."
		$TextboxResults.text = Get-RecipientPermission $MailboxUserSendAsPermission | Where-Object { ($_.IsInherited -eq $False) -and -not ($_.Trustee -like "NT AUTHORITY\SELF") } | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$displaySendOnBehalfPermissionsForMailboxToolStripMenuItem_Click = {
	$MailboxUserSendonPermission = Read-Host "Enter the UPN of the user you want to view Send On Behalf Of permission for"
	try
	{
		$TextboxResults.Text = "Getting Send On Behalf permissions for $MailboxUserSendonPermission..."
		$TextboxResults.text = Get-RecipientPermission $MailboxUserSendonPermission | Where-Object { ($_.IsInherited -eq $False) -and -not ($_.Trustee -like "NT AUTHORITY\SELF") } | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$removeFullAccessPermissionsForAMailboxToolStripMenuItem_Click = {
	$UserRemoveFullAccessRights = Read-Host "What user mailbox would you like modify Full Access rights to"
	$RemoveFullAccessRightsUser = Read-Host "Enter the UPN of the user you want to remove"
	try
	{
		$TextboxResults.Text = "Removing Full Access Permissions for $RemoveFullAccessRightsUser on account $UserRemoveFullAccessRights..."
		Remove-MailboxPermission  $UserRemoveFullAccessRights -User $RemoveFullAccessRightsUser -AccessRights FullAccess -Confirm:$False -ea 1
		$TextboxResults.text = Get-MailboxPermission $UserRemoveFullAccessRights | Where-Object { ($_.IsInherited -eq $False) -and -not ($_.User -like "NT AUTHORITY\SELF") } | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$revokeSendAsPermissionsForAMailboxToolStripMenuItem_Click = {
	$UserDeleteSendAsAccessOn = Read-Host "What user mailbox would you like to modify Send As permission for?"
	$UserDeleteSendAsAccess = Read-Host "Enter the UPN of the user you want to remove Send As access to?"
	try
	{
		$TextboxResults.Text = "Removing Send As permission for $UserDeleteSendAsAccess on account $UserDeleteSendAsAccessOn..."
		$TextboxResults.Text = Remove-RecipientPermission $UserDeleteSendAsAccessOn -AccessRights SendAs -Trustee $UserDeleteSendAsAccess -Confirm:$False | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$viewAllMailboxesAUserHasFullAccessToToolStripMenuItem_Click = {
	$ViewAllFullAccess = Read-Host "Enter the UPN of the account you want to view"
	try
	{
		$TextboxResults.Text = "Getting all mailboxes $ViewAllFullAccess has Full Access permissions to..."
		$TextboxResults.Text = Get-Mailbox | Get-MailboxPermission -User $ViewAllFullAccess | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$viewAllMailboxesAUserHasSendAsPermissionsToToolStripMenuItem_Click = {
	$ViewSendAsAccess = Read-Host "Enter the UPN of the account you want to view"
	try
	{
		$TextboxResults.Text = "Getting all mailboxes $ViewSendAsAccess has Send As permissions to..."
		$TextboxResults.Text = Get-Mailbox | Get-RecipientPermission -Trustee $ViewSendAsAccess | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$viewAllMailboxesAUserHasSendOnBehaldPermissionsToToolStripMenuItem_Click = {
	$ViewSendonBehalf = Read-Host "Enter the Name of the account you want to view"
	try
	{
		$TextboxResults.Text = "Getting all mailboxes $ViewSendonBehalf has Send On Behalf permissions to..."
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.GrantSendOnBehalfTo -match $ViewSendonBehalf } | Format-List DisplayName, GrantSendOnBehalfTo, PrimarySmtpAddress, RecipientType | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$removeAllPermissionsToAMailboxToolStripMenuItem_Click = {
	$UserDeleteAllAccessOn = Read-Host "What user mailbox would you like to modify permissions for?"
	$UserDeleteAllAccess = Read-Host "Enter the UPN of the user you want to remove access to?"
	try
	{
		$TextboxResults.Text = "Removing all permissions for $UserDeleteAllAccess on account $UserDeleteAllAccessOn..."
		$TextboxResults.Text = Remove-MailboxPermission -Identity $UserDeleteAllAccessOn -User $UserDeleteAllAccess -AccessRights FullAccess -InheritanceType All
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Forwarding

$getAllUsersForwardinToInternalRecipientToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all users forwarding to internal users..."
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.ForwardingAddress -ne $Null -and $_.RecipientType -eq "UserMailbox" } | Format-List Name, ForwardingAddress, DeliverToMailboxAndForward | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$forwardToInternalRecipientAndDontSaveLocalCopyToolStripMenuItem_Click = {
	$UsertoFWD2 = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	$Fwd2me2 = Read-Host "Enter the Name, Display Name, Alias, or Email Address of the user to forward to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding from $UsertoFWD2 to $Fwd2me2..."
		Set-Mailbox  $UsertoFWD2 -ForwardingAddress $Fwd2me2 -DeliverToMailboxAndForward $False
		$TextboxResults.Text = Get-Mailbox $UsertoFWD2 | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$forwardToExternalAddressAndSaveLocalCopyToolStripMenuItem_Click = {
	$UsertoFWD3 = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	$Fwd2me2External = Read-Host "Enter the external email address to forward to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding from $UsertoFWD3 to $Fwd2me2External..."
		Set-Mailbox $UsertoFWD3 -ForwardingsmtpAddress $Fwd2me2External -DeliverToMailboxAndForward $true
		$TextboxResults.Text = Get-Mailbox $UsertoFWD3 | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Format-List | Out-String
		
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$forwardToExternalAddressAndDontSaveLocalCopyToolStripMenuItem_Click = {
	$UsertoFWD4 = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	$Fwd2me2External2 = Read-Host "Enter the external email address to forward to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding from $UsertoFWD4 to $Fwd2me2External2..."
		Set-Mailbox $UsertoFWD4 -ForwardingsmtpAddress $Fwd2me2External2 -DeliverToMailboxAndForward $False
		$TextboxResults.Text = Get-Mailbox $UsertoFWD4 | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getForwardingInfoForAUserToolStripMenuItem_Click = {
	$UserFwdInfo = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user"
	Try
	{
		$TextboxResults.Text = "Getting forwarding info for $UserFwdInfo..."
		$TextboxResults.Text = Get-Mailbox $UserFwdInfo | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$removeExternalForwadingForAUserToolStripMenuItem_Click = {
	$RemoveFWDfromUserExternal = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address"
	Try
	{
		$TextboxResults.Text = "Removing all external forwarding from $RemoveFWDfromUserExternal..."
		Set-Mailbox $RemoveFWDfromUserExternal -ForwardingSmtpAddress $Null
		$TextboxResults.Text = Get-Mailbox $RemoveFWDfromUserExternal | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$removeAllForwardingForAUserToolStripMenuItem_Click = {
	$RemoveAllFWDforUser = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address"
	Try
	{
		$TextboxResults.Text = "Removing all forwarding from $RemoveAllFWDforUser..."
		Set-Mailbox $RemoveAllFWDforUser -ForwardingAddress $Null -ForwardingSmtpAddress $Null
		$TextboxResults.Text = Get-Mailbox $RemoveAllFWDforUser | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$removeInternalForwardingForUserToolStripMenuItem_Click = {
	$RemoveFWDfromUser = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address"
	Try
	{
		$TextboxResults.Text = "Removing all internal forwarding from $RemoveFWDfromUser..."
		Set-Mailbox $RemoveFWDfromUser -ForwardingAddress $Null
		$TextboxResults.Text = Get-Mailbox $RemoveFWDfromUser | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$forwardToInternalRecipientAndSaveLocalCopyToolStripMenuItem_Click = {
	$UsertoFWD = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	$Fwd2me = Read-Host "Enter the Name, Display Name, Alias, or Email Address of the user to forward to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding from $UsertoFWD to $Fwd2me..."
		Set-Mailbox  $UsertoFWD -ForwardingAddress $Fwd2me -DeliverToMailboxAndForward $True
		$TextboxResults.Text = Get-Mailbox $UsertoFWD | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getAllUsersForwardingToExternalRecipientToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all users forwarding to external users..."
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.ForwardingSmtpAddress -ne $Null -and $_.RecipientType -eq "UserMailbox" } | Format-List Name, ForwardingSmtpAddress, DeliverToMailboxAndForward | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
	
}

$removeAllForwardingForAllUsersToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Removing all forwarding from all users..."
		$AllMailboxes = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingAddress $Null -ForwardingSmtpAddress $Null
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$removeExternalForwardingForAllUsersToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Removing all external forwarding from all users..."
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingSmtpAddress $Null
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Format-List Name, DeliverToMailboxAndForward, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$removeInternalForwardingForAllUsersToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Removing all internal forwarding from all users..."
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingAddress $Null
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$forwardAllUsersEmailToExternalRecipientAndSaveALocalCopyToolStripMenuItem_Click = {
	$ForwardAllToExternal = Read-Host "Enter the email to forward all email to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding for all users to $ForwardAllToExternal..."
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingsmtpAddress $ForwardAllToExternal -DeliverToMailboxAndForward $True
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Format-List Name, DeliverToMailboxAndForward, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$forwardAllUsersEmailToExternalRecipientAndDontSaveALocalCopyToolStripMenuItem_Click = {
	$ForwardAllToExternal2 = Read-Host "Enter the email to forward all email to"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding for all users to $ForwardAllToExternal2..."
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingsmtpAddress $ForwardAllToExternal2 -DeliverToMailboxAndForward $False
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Format-List Name, DeliverToMailboxAndForward, ForwardingSmtpAddress | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$forwardAllUsersEmailToInternalRecipientAndSaveLocalCopyToolStripMenuItem_Click = {
	$ForwardAllToInternal = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding for all users to $ForwardAllToInternal..."
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingAddress $ForwardAllToInternal -DeliverToMailboxAndForward $True
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$forwardAllUsersEmailToInternalRecipientAndDontSaveLocalCopyToolStripMenuItem_Click = {
	$ForwardAllToInternal2 = Read-Host "Enter the users UPN, Display Name, Alias, or Email Address of the user to forward"
	Try
	{
		$TextboxResults.Text = "Setting up forwarding for all users to $ForwardAllToInternal2..."
		Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Set-Mailbox -ForwardingAddress $ForwardAllToInternal2 -DeliverToMailboxAndForward $True
		$TextboxResults.Text = Get-Mailbox | Where-Object { $_.RecipientType -eq "UserMailbox" } | Format-List Name, DeliverToMailboxAndForward, ForwardingAddress | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}



###GROUPS###

	#Distribution Groups

$displayDistributionGroupsToolStripMenuItem_Click={
	try
	{
		$TextboxResults.Text = "Getting all Distribution Groups..."
		$TextboxResults.text = Get-DistributionGroup | Where-Object { $_.GroupType -notlike "Universal, SecurityEnabled"} | Format-List DisplayName, SamAccountName, GroupType, IsDirSynced, EmailAddresses | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$createADistributionGroupToolStripMenuItem_Click = {
	$NewDistroGroup = Read-Host "What is the name of the new Distribution Group?"
	try
	{
		$TextboxResults.Text = "Creating the $NewDistroGroup Distribution Group..."
		$TextboxResults.Text = New-DistributionGroup -Name $NewDistroGroup | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$deleteADistributionGroupToolStripMenuItem_Click = {
	$DeleteDistroGroup = Read-Host "Enter the name of the Distribtuion group you want deleted."
	try
	{
		$TextboxResults.Text = "Deleting the $DeleteDistroGroup Distribution Group..."
		Remove-DistributionGroup $DeleteDistroGroup
		$TextboxResults.Text = "Getting list of distribution groups"
		$TextboxResults.text = Get-DistributionGroup | Where-Object { $_.GroupType -notlike "Universal, SecurityEnabled" } | Format-List DisplayName | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$allowDistributionGroupToReceiveExternalEmailToolStripMenuItem_Click = {
	$AllowExternalEmail = Read-Host "Enter the name of the Distribtuion Group you want to allow external email to"
	try
	{
		$TextboxResults.Text = "Allowing extneral senders for the $AllowExternalEmail Distribution Group..."
		Set-DistributionGroup $AllowExternalEmail -RequireSenderAuthenticationEnabled $False 
		$TextboxResults.text = Get-DistributionGroup $AllowExternalEmail | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$hideDistributionGroupFromGALToolStripMenuItem_Click = {
	$GroupHideGAL = Read-Host "Enter the name of the Distribtuion Group you want to allow external email to"
	try
	{
		$TextboxResults.Text = "Hiding the $GroupHideGAL from the Global Address List..."
		Set-DistributionGroup $GroupHideGAL -HiddenFromAddressListsEnabled $True
		$TextboxResults.text = Get-DistributionGroup $GroupHideGAL | Format-List Name, HiddenFromAddressListsEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$displayDistributionGroupMembersToolStripMenuItem_Click = {
	$ListDistributionGroupMembers = Read-Host "Enter the name of the Distribution Group you want to list members of"
	try
	{
		$TextboxResults.Text = "Getting all members of the $ListDistributionGroupMembers Distrubution Group..."
		$TextboxResults.Text = Get-DistributionGroupMember $ListDistributionGroupMembers | Format-List DisplayName | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$addAUserToADistributionGroupToolStripMenuItem_Click = {
	$DistroGroupAdd = Read-Host "Enter the name of the Distribution Group"
	$DistroGroupAddUser = Read-Host "Enter the UPN of the user you wish to add to $DistroGroupAdd"
	try
	{
		$TextboxResults.Text = "Adding $DistroGroupAddUser to the $DistroGroupAdd Distribution Group..."
		Add-DistributionGroupMember -Identity $DistroGroupAdd -Member $DistroGroupAddUser
		$TextboxResults.Text = "Getting members of $DistroGroupAdd..."
		$TextboxResults.Text = Get-DistributionGroupMember $DistroGroupAdd | Format-List DisplayName | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$removeAUserADistributionGroupToolStripMenuItem_Click = {
	$DistroGroupRemove = Read-Host "Enter the name of the Distribution Group"
	$DistroGroupRemoveUser = Read-Host "Enter the UPN of the user you wish to remove from $DistroGroupRemove"
	try
	{
		$TextboxResults.Text = "Removing $DistroGroupRemoveUser from the $DistroGroupRemove Distribution Group..."
		Remove-DistributionGroupMember -Identity $DistroGroupRemove -Member $DistroGroupRemoveUser
		$TextboxResults.Text = Get-DistributionGroupMember $DistroGroupRemove | Format-List DisplayName | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$addAllUsersToADistributionGroupToolStripMenuItem_Click = {
		$users = Get-Mailbox | Select-Object -ExpandProperty Alias
		$AddAllUsersToSingleDistro = Read-Host "Enter the name of the Distribution Group you want to add all users to"
		try
		{
		$TextboxResults.Text = "Adding all users to the $AddAllUsersToSingleDistro distribution group..."
			Foreach ($user in $users) { Add-DistributionGroupMember -Identity $AddAllUsersToSingleDistro -Member $user }
			$TextboxResults.Text = Get-DistributionGroupMember $AddAllUsersToSingleDistro | Format-List DisplayName | Out-String
		}
		catch
		{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
		}
}

$getDetailedInfoForDistributionGroupToolStripMenuItem_Click = {
	$DetailedInfoMailDistroGroup = Read-Host "Enter the group name"
	Try
	{
		$TextboxResults.Text = "Getting detailed info about the $DetailedInfoMailDistroGroup group..."
		$TextboxResults.text = Get-DistributionGroup $DetailedInfoMailDistroGroup | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$allowAllDistributionGroupsToReceiveExternalEmailToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Allowing extneral senders for all Distribution Groups..."
		Get-DistributionGroup | Set-DistributionGroup -RequireSenderAuthenticationEnabled $False
		$TextboxResults.text = Get-DistributionGroup | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$denyDistributionGroupFromReceivingExternalEmailToolStripMenuItem_Click = {
	$DenyExternalEmail = Read-Host "Enter the name of the Distribtuion Group you want to deny external email to"
	try
	{
		$TextboxResults.Text = "Denying extneral senders for the $DenyExternalEmail Distribution Group..."
		Set-DistributionGroup $DenyExternalEmail -RequireSenderAuthenticationEnabled $True
		$TextboxResults.text = Get-DistributionGroup $DenyExternalEmail | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$denyAllDistributionGroupsFromReceivingExternalEmailToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Denying extneral senders for all Distribution Groups..."
		Get-DistributionGroup | Set-DistributionGroup -RequireSenderAuthenticationEnabled $True
		$TextboxResults.text = Get-DistributionGroup | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getExternalEmailStatusForADistributionGroupToolStripMenuItem_Click = {
	$ExternalEmailStatus = Read-Host "Enter the Distribution Group"
	try
	{
		$TextboxResults.Text = "Getting external email status for $ExternalEmailStatus..."
		$TextboxResults.text = Get-DistributionGroup $ExternalEmailStatus | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getExternalEmailStatusForAllDistributionGroupsToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting external email status for all distribution groups..."
		$TextboxResults.text = Get-DistributionGroup | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

	#Unified Groups

$getListOfUnifiedGroupsToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting list of all unified groups..."
		$TextboxResults.Text = Get-UnifiedGroup | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$listMembersOfAGroupToolStripMenuItem_Click = {
	$GetUnifiedGroupMembers = Read-Host "Enter the name of the group you want to view members for."
	try
	{
		$TextboxResults.Text = "Getting all members of the $GetUnifiedGroupMembers group..."
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $GetUnifiedGroupMembers –LinkType Members | Format-List DisplayName | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$removeAGroupToolStripMenuItem_Click = {
	$RemoveUnifiedGroup = Read-Host "Enter the name of the group you want to remove"
	try
	{
		$TextboxResults.Text = "Removing the $RemoveUnifiedGroup group..."
		Remove-UnifiedGroup $RemoveUnifiedGroup
		$TextboxResults.Text = Get-UnifiedGroup | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
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
		Add-UnifiedGroupLinks $UnifiedGroupAddUser –Links $UnifiedGroupUserAdd –LinkType $UnifiedGroupAccess
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $UnifiedGroupAddUser –LinkType Members | Format-List DisplayName | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$createANewGroupToolStripMenuItem_Click = {
	$NewUnifiedGroupName = Read-Host "Enter the Display Name of the new group"
	$NewUnifiedGroupAccessType = Read-Host "Enter the Access Type for the group $NewUnifiedGroupName (Public or Private)"
	try
	{
		$TextboxResults.Text = "Creating a the $NewUnifiedGroupName group..."
		New-UnifiedGroup –DisplayName $NewUnifiedGroupName -AccessType $NewUnifiedGroupAccessType
		$TextboxResults.Text = Get-UnifiedGroup $NewUnifiedGroupName | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$listOwnersOfAGroupToolStripMenuItem_Click = {
	$GetUnifiedGroupOwners = Read-Host "Enter the name of the group you want to view members for."
	try
	{
		$TextboxResults.Text = "Getting all owners of the $GetUnifiedGroupOwners group..."
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $GetUnifiedGroupOwners –LinkType Owners | Format-List DisplayName | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$listSubscribersOfAGroupToolStripMenuItem_Click = {
	$GetUnifiedGroupSubscribers = Read-Host "Enter the name of the group you want to view members for."
	try
	{
		$TextboxResults.Text = "Getting all subscribers of the $GetUnifiedGroupSubscribers group..."
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $GetUnifiedGroupSubscribers –LinkType Subscribers | Format-List DisplayName | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$addAnOwnerToAGroupToolStripMenuItem_Click = {
	$TextboxResults.Text = "Important! The user must be a member of the group prior to becoming an owner"
	$UnifiedGroupAddOwner = Read-Host "Enter the name of the group you want to modify ownership for"
	$AddUnifiedGroupOwner = Read-Host "Enter the UPN of the user you want to become an owner"
	try
	{
		$TextboxResults.Text = "Adding $AddUnifiedGroupOwner as an owner of the $UnifiedGroupAddOwner group..."
		Add-UnifiedGroupLinks -Identity $UnifiedGroupAddOwner -LinkType Owners -Links $AddUnifiedGroupOwner
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $UnifiedGroupAddOwner –LinkType Owners | Format-List DisplayName | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$addASubscriberToAGroupToolStripMenuItem_Click = {
	$UnifiedGroupAddSubscriber = Read-Host "Enter the name of the group you want to add a subscriber for"
	$AddUnifiedGroupSubscriber = Read-Host "Enter the UPN of the user you want to add as a subscriber"
	try
	{
		$TextboxResults.Text = "Adding $AddUnifiedGroupSubscriber as a subscriber to the $UnifiedGroupAddSubscriber group..."
		Add-UnifiedGroupLinks -Identity $UnifiedGroupAddSubscriber -LinkType Owners -Links $AddUnifiedGroupSubscriber
		$TextboxResults.Text = Get-UnifiedGroupLinks –Identity $UnifiedGroupAddSubscriber –LinkType Subscribers | Format-List DisplayName | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

	#Security Groups

$createARegularSecurityGroupToolStripMenuItem_Click = {
	$SecurityGroupName = Read-Host "Enter a name for the new Security Group"
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Creating the $SecurityGroupName security group..."
		$TextboxResults.Text = New-MsolGroup -DisplayName $SecurityGroupName | Format-List | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Creating the $SecurityGroupName security group..."
		$TextboxResults.Text = New-MsolGroup -DisplayName $SecurityGroupName -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getAllRegularSecurityGroupsToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting list of all Security groups..."
		$TextboxResults.Text = Get-MsolGroup -GroupType Security | Format-List DisplayName, GroupType | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting list of all Security groups..."
		$TextboxResults.Text = Get-MsolGroup -TenantId $PartnerComboBox.SelectedItem.TenantID -GroupType Security | Format-List DisplayName, GroupType | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getAllMailEnabledSecurityGroupsToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all Mail Enabled Security Groups..."
		$TextboxResults.text = Get-DistributionGroup | Where-Object { $_.GroupType -like "Universal, SecurityEnabled" } | Format-List DisplayName, SamAccountName, GroupType, IsDirSynced, EmailAddresses | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$createAMailEnabledSecurityGroupToolStripMenuItem_Click = {
	$MailEnabledSecurityGroup = Read-Host "Enter the name of the security group"
	$MailEnabledSMTPAddress = Read-Host "Enter the primary SMTP address for $MailEnabledSecurityGroup"
	Try
	{
		$TextboxResults.Text = "Creating the $MailEnabledSecurityGroup security group..."
		$TextboxResults.Text = New-DistributionGroup -Name $MailEnabledSecurityGroup -Type Security -PrimarySmtpAddress $MailEnabledSMTPAddress | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$addAUserToAMailEnabledSecurityGroupToolStripMenuItem_Click = {
	$MailEnabledGroupAdd = Read-Host "Enter the name of the Group"
	$MailEnabledGroupAddUser = Read-Host "Enter the UPN of the user you wish to add to $MailEnabledGroupAdd"
	try
	{
		$TextboxResults.Text = "Adding $MailEnabledGroupAddUser to the $MailEnabledGroupAdd Group..."
		Add-DistributionGroupMember -Identity $MailEnabledGroupAdd -Member $MailEnabledGroupAddUser
		$TextboxResults.Text = "Getting members of $MailEnabledGroupAdd..."
		$TextboxResults.Text = Get-DistributionGroupMember $MailEnabledGroupAdd | Format-List Displayname | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$allowSecurityGroupToRecieveExternalMailToolStripMenuItem_Click = {
	$AllowExternalEmailSecurity = Read-Host "Enter the name of the Group you want to allow external email to"
	try
	{
		$TextboxResults.Text = "Allowing extneral senders for the $AllowExternalEmailSecurity Group..."
		Set-DistributionGroup $AllowExternalEmailSecurity -RequireSenderAuthenticationEnabled $False
		$TextboxResults.text = Get-DistributionGroup $AllowExternalEmailSecurity | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getDetailedInfoForMailEnabledSecurityGroupToolStripMenuItem_Click = {
	$DetailedInfoMailEnabledSecurityGroup = Read-Host "Enter the group name"
	Try
	{
		$TextboxResults.Text = "Getting detailed info about the $DetailedInfoMailEnabledSecurityGroup group..."
		$TextboxResults.text = Get-DistributionGroup $DetailedInfoMailEnabledSecurityGroup | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$removeMailENabledSecurityGroupToolStripMenuItem_Click = {
	$DeleteMailEnabledSecurityGroup = Read-Host "Enter the name of the group you want deleted."
	try
	{
		$TextboxResults.Text = "Deleting the $DeleteMailEnabledSecurityGroup Group..."
		Remove-DistributionGroup $DeleteMailEnabledSecurityGroup
		$TextboxResults.Text = "Getting list of mail enabled security groups..."
		$TextboxResults.text = Get-DistributionGroup | Where-Object { $_.GroupType -like "Universal, SecurityEnabled" } | Format-List DisplayName | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$denySecurityGroupFromRecievingExternalEmailToolStripMenuItem_Click = {
	$DenyExternalEmailSecurity = Read-Host "Enter the name of the Group you want to deny external email to"
	try
	{
		$TextboxResults.Text = "Denying extneral senders for the $DenyExternalEmailSecurity Group..."
		Set-DistributionGroup $DenyExternalEmailSecurity -RequireSenderAuthenticationEnabled $True
		$TextboxResults.text = Get-DistributionGroup $DenyExternalEmailSecurity | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}



###RESOURCE MAILBOX###

	#Booking Options

$allowConflictMeetingsToolStripMenuItem_Click = {
	$ConflictMeetingAllow = Read-Host "Enter the Room Name of the Resource Calendar you want to allow conflicts"
	try
	{
		$TextboxResults.Text = "Allowing conflict meetings $ConflictMeetingAllow..."
		Set-CalendarProcessing $ConflictMeetingAllow -AllowConflicts $True
		$TextboxResults.Text = Get-CalendarProcessing -identity $ConflictMeetingAllow | Format-List Identity, AllowConflicts | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$denyConflictMeetingsForAllResourceMailboxesToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Denying conflict meeting for all room calendars..."
		Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AllowConflicts $False
		$TextboxResults.Text = Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Get-CalendarProcessing | Format-List Identity, AllowConflicts | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$allowConflicMeetingsForAllResourceMailboxesToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Allowing conflict meeting for all room calendars..."
		Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AllowConflicts $True
		$TextboxResults.Text = Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Get-CalendarProcessing | Format-List Identity, AllowConflicts | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$disallowconflictmeetingsToolStripMenuItem_Click = {
	$ConflictMeetingDeny = Read-Host "Enter the Room Name of the Resource Calendar you want to disallow conflicts"
	try
	{
		$TextboxResults.Text = "Denying conflict meetings for $ConflictMeetingDeny..."
		Set-CalendarProcessing $ConflictMeetingDeny -AllowConflicts $False
		$TextboxResults.Text = Get-CalendarProcessing -identity $ConflictMeetingDeny | Format-List Identity, AllowConflicts | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$enableAutomaticBookingForAllResourceMailboxToolStripMenuItem1_Click = {
		Try
		{
			$TextboxResults.Text = "Enabling automatic booking on all room calendars..."
			Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AutomateProcessing:AutoAccept
			$TextboxResults.Text = Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Get-CalendarProcessing | Format-List Identity, AutomateProcessing | Out-String
		}
		Catch
		{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
		}
	
}

$GetRoomMailBoxCalendarProcessingToolStripMenuItem_Click = {
	$RoomMailboxCalProcessing = Read-Host "Enter the Calendar Name you want to view calendar processing information for"
	try
	{
		$TextboxResults.Text = "Getting calendar processing information for $RoomMailboxCalProcessing..."
		$TextboxResults.Text = Get-Mailbox $RoomMailboxCalProcessing | Get-CalendarProcessing | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

	#Room Mailbox

$convertAMailboxToARoomMailboxToolStripMenuItem_Click = {
	$MailboxtoRoom = Read-Host "What user would you like to convert to a Room Mailbox? Please enter the full email address"
	Try
	{
		$TextboxResults.Text = "Converting $MailboxtoRoom to a Room Mailbox..."
		Set-Mailbox $MailboxtoRoom -Type Room
		$TextboxResults.Text = Get-MailBox $MailboxtoRoom | Format-List Name, ResourceType, PrimarySmtpAddress, EmailAddresses, UserPrincipalName, IsMailboxEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$createANewRoomMailboxToolStripMenuItem_Click = {
	$NewRoomMailbox = Read-Host "Enter the name of the new room mailbox"
	Try
	{
		$TextboxResults.Text = "Creating the $NewRoomMailbox Room Mailbox..."
		$TextboxResults.Text = New-Mailbox -Name $NewRoomMailbox -Room | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getListOfRoomMailboxesToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting list of all Room Mailboxes..."
		$TextboxResults.Text = Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Format-List Identity, PrimarySmtpAddress, EmailAddresses, UserPrincipalName | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$removeARoomMailboxToolStripMenuItem_Click = {
	$RemoveRoomMailbox = Read-Host "Enter the name of the room mailbox"
	Try
	{
		$TextboxResults.Text = "Removing the $RemoveRoomMailbox Room Mailbox..."
		Remove-Mailbox $RemoveRoomMailbox
		$TextboxResults.Text = Get-MailBox | Where-Object { $_.ResourceType -eq "Room" } | Format-Table -AutoSize | Out-String
		
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}




###JUNK EMAIL CONFIGURATION###

	#Blacklist

$blacklistDomainForAllToolStripMenuItem_Click = {
	$BlacklistDomain = Read-Host "Enter the domain you want to blacklist for all users. EX: google.com"
	try
	{
		$TextboxResults.Text = "Blacklisting $BlacklistDomain for all users..."
		Get-Mailbox | Set-MailboxJunkEmailConfiguration -BlockedSendersAndDomains $BlacklistDomain
		$TextboxResults.Text = Get-Mailbox | Get-MailboxJunkEmailConfiguration | Format-List Identity, BlockedSendersAndDomains, Enabled | Out-String
		
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$blacklistDomainForASingleUserToolStripMenuItem_Click = {
	$Blockeddomainuser = Read-Host "Enter the UPN of the user you want to modify junk email for"
	$BlockedDomain2 = Read-Host "Enter the domain you want to blacklist"
	try
	{
		$TextboxResults.Text = "Blacklisting $BlockedDomain2 for $Blockeddomainuser..."
		Set-MailboxJunkEmailConfiguration -Identity $Blockeddomainuser -BlockedSendersAndDomains $BlockedDomain2
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $Blockeddomainuser | Format-List Identity, BlockedSendersAndDomains | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$blacklistASpecificEmailAddressForAllToolStripMenuItem_Click = {
	$BlockSpecificEmailForAll = Read-Host "Enter the email address you want to blacklist for all"
	try
	{
		$TextboxResults.Text = "Blacklisting $BlockSpecificEmailForAll for all users..."
		Get-Mailbox | Set-MailboxJunkEmailConfiguration -BlockedSendersAndDomains $BlockSpecificEmailForAll
		$TextboxResults.Text = Get-Mailbox | Get-MailboxJunkEmailConfiguration | Format-List Identity, BlockedSendersAndDomains, Enabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$blacklistASpecificEmailAddressForASingleUserToolStripMenuItem_Click = {
	$ModifyblacklistforaUser = Read-Host "Enter the user you want to modify the blacklist for"
	$DenySpecificEmailForOne = Read-Host "Enter the email address you want to whitelist for a single user"
	try
	{
		$TextboxResults.Text = "Blacklisting $DenySpecificEmailForOne for $ModifyblacklistforaUser..."
		Set-MailboxJunkEmailConfiguration -Identity $ModifyblacklistforaUser -BlockedSendersAndDomains $DenySpecificEmailForOne
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $ModifyblacklistforaUser | Format-List Identity, BlockedSendersAndDomains, Enabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Junk Email General Items

$checkSafeAndBlockedSendersForAUserToolStripMenuItem_Click = {
	$CheckSpamUser = Read-Host "Enter the UPN of the user you want to check blocked and allowed senders for"
	try
	{
		$TextboxResults.Text = "Getting safe and blocked senders for $CheckSpamUser..."
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $CheckSpamUser | Format-List Identity, TrustedListsOnly, ContactsTrusted, TrustedSendersAndDomains, BlockedSendersAndDomains, TrustedRecipientsAndDomains, IsValid | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Whitelist

$whitelistDomainForAllToolStripMenuItem_Click = {
	$AllowedDomain = Read-Host "Enter the domain you want to whitelist for all users. EX: google.com"
	try
	{
		$TextboxResults.Text = "Whitelisting $AllowedDomain for all..."
		Get-Mailbox | Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains $AllowedDomain
		$TextboxResults.Text = Get-Mailbox | Get-MailboxJunkEmailConfiguration | Format-List Identity, TrustedSendersAndDomains, TrustedRecipientsAndDomains, Enabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$whitelistDomainForASingleUserToolStripMenuItem_Click = {
	$Alloweddomainuser = Read-Host "Enter the UPN of the user you want to modify junk email for"
	$AllowedDomain2 = Read-Host "Enter the domain you want to whitelist"
	try
	{
		$TextboxResults.Text = "Whitelisting $AllowedDomain2 for $Alloweddomainuser..."
		Set-MailboxJunkEmailConfiguration -Identity $Alloweddomainuser -TrustedSendersAndDomains $AllowedDomain2
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $Alloweddomainuser | Format-List Identity, TrustedSendersAndDomains, TrustedRecipientsAndDomains | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	} 
}

$whitelistASpecificEmailAddressForAllToolStripMenuItem_Click = {
	$AllowSpecificEmailForAll = Read-Host "Enter the email address you want to whitelist for all"
	try
	{
		$TextboxResults.Text = "Whitelisting $AllowSpecificEmailForAll for all..."
		Get-Mailbox | Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains $AllowSpecificEmailForAll
		$TextboxResults.Text = Get-Mailbox | Get-MailboxJunkEmailConfiguration | Format-List Identity, TrustedSendersAndDomains, TrustedRecipientsAndDomains, Enabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$whitelistASpecificEmailAddressForASingleUserToolStripMenuItem_Click = {
	$ModifyWhitelistforaUser = Read-Host "Enter the user you want to modify the whitelist for"
	$AllowSpecificEmailForOne = Read-Host "Enter the email address you want to whitelist for a single user"
	try
	{
		$TextboxResults.Text = "Whitelisting $AllowSpecificEmailForOne for $ModifyWhitelistforaUser..."
		Set-MailboxJunkEmailConfiguration -Identity $ModifyWhitelistforaUser -TrustedSendersAndDomains $AllowSpecificEmailForOne
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $ModifyWhitelistforaUser | Format-List Identity, TrustedSendersAndDomains, TrustedRecipientsAndDomains, Enabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}



###ADMIN###

	#OWA

$disableAccessToOWAForASingleUserToolStripMenuItem_Click = {
	$DisableOWAforUser = Read-Host "Enter the UPN of the user you want to disable OWA access for"
	try
	{
		$TextboxResults.Text = "Disabling OWA access for $DisableOWAforUser..."
		Set-CASMailbox $DisableOWAforUser -OWAEnabled $False
		$TextboxResults.Text = Get-CASMailbox $DisableOWAforUser | Format-List DisplayName, OWAEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$enableOWAAccessForASingleUserToolStripMenuItem_Click = {
	$EnableOWAforUser = Read-Host "Enter the UPN of the user you want to enable OWA access for"
	try
	{
		$TextboxResults.Text = "Enabling OWA access for $EnableOWAforUser..."
		Set-CASMailbox $EnableOWAforUser -OWAEnabled $True
		$TextboxResults.Text = Get-CASMailbox $EnableOWAforUser | Format-List DisplayName, OWAEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$disableOWAAccessForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Disabling OWA access for all..."
		Get-Mailbox | Set-CASMailbox -OWAEnabled $False
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox | Format-List DisplayName, OWAEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$enableOWAAccessForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Enabling OWA access for all..."
		Get-Mailbox | Set-CASMailbox -OWAEnabled $True
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox | Format-List DisplayName, OWAEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getOWAAccessForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting OWA info for all users..."
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox | Format-List DisplayName, OWAEnabled, OwaMailboxPolicy, OWAforDevicesEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getOWAInfoForASingleUserToolStripMenuItem_Click = {
	$OWAAccessUser = Read-Host "Enter the UPN for the User you want to view OWA info for"
	try
	{
		$TextboxResults.Text = "Getting OWA Access for $OWAAccessUser..."
		$TextboxResults.Text = Get-CASMailbox -identity $OWAAccessUser | Format-List DisplayName, OWAEnabled, OwaMailboxPolicy, OWAforDevicesEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#ActiveSync

$getActiveSyncDevicesForAUserToolStripMenuItem_Click = {
	$ActiveSyncDevicesUser = Read-Host "Enter the UPN of the user you wish to see ActiveSync Devices for"
	try
	{
		$TextboxResults.Text = "Getting ActiveSync device info for $ActiveSyncDevicesUser..."
		$TextboxResults.Text = Get-MobileDeviceStatistics -Mailbox $ActiveSyncDevicesUser | Format-List DeviceFriendlyName, DeviceModel, DeviceOS, DeviceMobileOperator, DeviceType, Status, FirstSyncTime, LastPolicyUpdateTime, LastSyncAttemptTime, LastSuccessSync, LastPingHeartbeat, DeviceAccessState, IsValid  | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$disableActiveSyncForAUserToolStripMenuItem_Click = {
	$DisableActiveSyncForUser = Read-Host "Enter the UPN of the user you wish to disable ActiveSync for"
	try
	{
		$TextboxResults.Text = "Disabling ActiveSync for $DisableActiveSyncForUser..."
		Set-CASMailbox $DisableActiveSyncForUser -ActiveSyncEnabled $False 
		$TextboxResults.Text = Get-CASMailbox -Identity $DisableActiveSyncForUser | Format-List DisplayName, ActiveSyncEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$enableActiveSyncForAUserToolStripMenuItem_Click = {
	$EnableActiveSyncForUser = Read-Host "Enter the UPN of the user you wish to enable ActiveSync for"
	try
	{
		$TextboxResults.Text = "Enabling ActiveSync for $EnableActiveSyncForUser... "
		Set-CASMailbox $EnableActiveSyncForUser -ActiveSyncEnabled $True
		$TextboxResults.Text = Get-CASMailbox -Identity $EnableActiveSyncForUser | Format-List DisplayName, ActiveSyncEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$viewActiveSyncInfoForAUserToolStripMenuItem_Click = {
	$ActiveSyncInfoForUser = Read-Host "Enter the UPN for the user you want to view ActiveSync info for"
	try
	{
		$TextboxResults.Text = "Getting ActiveSync info for $ActiveSyncInfoForUser..."
		$TextboxResults.Text = Get-CASMailbox -Identity $ActiveSyncInfoForUser | Format-List DisplayName, ActiveSyncEnabled, ActiveSyncAllowedDeviceIDs, ActiveSyncBlockedDeviceIDs, ActiveSyncMailboxPolicy, ActiveSyncMailboxPolicyIsDefaulted, ActiveSyncDebugLogging, HasActiveSyncDevicePartnership | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$disableActiveSyncForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Disabling ActiveSync for all..."
		Get-Mailbox | Set-CASMailbox -ActiveSyncEnabled $False
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox  | Format-List DisplayName, ActiveSyncEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getActiveSyncInfoForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting ActiveSync info for all..."
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox | Format-List DisplayName, ActiveSyncEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
		
	}
}

$enableActiveSyncForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Enabling ActiveSync for all.."
		Get-Mailbox | Set-CASMailbox -ActiveSyncEnabled $True
		$TextboxResults.Text = Get-Mailbox | Get-CASMailbox | Format-List DisplayName, ActiveSyncEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#PowerShell

$disableAccessToPowerShellForAUserToolStripMenuItem_Click = {
	$DisablePowerShellforUser = Read-Host "Enter the UPN of the user you want to disable PowerShell access for"
	try
	{
		$TextboxResults.Text = "Disabling PowerShell access for $DisablePowerShellforUser..."
		Set-User $DisablePowerShellforUser -RemotePowerShellEnabled $False
		$TextboxResults.Text = Get-User $DisablePowerShellforUser | Format-List DisplayName, RemotePowerShellEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$displayPowerShellRemotingStatusForAUserToolStripMenuItem_Click = {
	$PowerShellRemotingStatusUser = Read-Host "Enter the UPN of the user you want to view PowerShell Remoting for"
	try
	{
		$TextboxResults.Text = "Getting PowerShell info for $PowerShellRemotingStatusUser..."
		$TextboxResults.Text = Get-User $PowerShellRemotingStatusUser | Format-List DisplayName, RemotePowerShellEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$enableAccessToPowerShellForAUserToolStripMenuItem_Click = {
	$EnablePowerShellforUser = Read-Host "Enter the UPN of the user you want to enable PowerShell access for"
	try
	{
		$TextboxResults.Text = "Enabling PowerShell access for $EnablePowerShellforUser..."
		Set-User $EnablePowerShellforUser -RemotePowerShellEnabled $True
		$TextboxResults.Text = Get-User $EnablePowerShellforUser | Format-List DisplayName, RemotePowerShellEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Message Trace

$messageTraceToolStripMenuItem_Click = {
	
}

$GetAllRecentToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting recent messages..."
		$TextboxResults.Text = Get-MessageTrace | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$fromACertainSenderToolStripMenuItem1_Click = {
	$MessageTraceSender = Read-Host "Enter the senders email address"
	try
	{
		$TextboxResults.Text = "Getting messages from $MessageTraceSender..."
		$TextboxResults.Text = Get-MessageTrace -SenderAddress $MessageTraceSender | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$toACertainRecipientToolStripMenuItem_Click = {
	$MessageTraceRecipient = Read-Host "Enter the recipients email address"
	try
	{
		$TextboxResults.Text = "Getting messages sent to $MessageTraceRecipient..."
		$TextboxResults.Text = Get-MessageTrace -RecipientAddress $MessageTraceRecipient | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getFailedMessagesToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting failed messages..."
		$TextboxResults.Text = Get-MessageTrace -Status "Failed" | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$GetMessagesBetweenDatesToolStripMenuItem_Click = {
	$MessageTraceStartDate = Read-Host "Enter the start date. EX: 06/13/2015 or 09/01/2015 5:00 PM"
	$MessageTraceEndDate = Read-Host "Enter the end date. EX: 06/15/2015 or 09/01/2015 5:00 PM"
	try
	{
		$TextboxResults.Text = "Getting messages between $MessageTraceStartDate and $MessageTraceEndDate..."
		$TextboxResults.Text = Get-MessageTrace -StartDate $MessageTraceStartDate -EndDate $MessageTraceEndDate | Format-List Received, SenderAddress, RecipientAddress, FromIP, ToIP, Subject, Size, Status | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Company Info

$getTechnicalNotificationEmailToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting technical account info..."
		$TextboxResults.Text = Get-MsolCompanyInformation | Format-List TechnicalNotificationEmails | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting technical account info..."
		$TextboxResults.Text = Get-MsolCompanyInformation -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List TechnicalNotificationEmails | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$lastDirSyncTimeToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last DirSync time..."
		$TextboxResults.Text = Get-MsolCompanyInformation | Format-List LastDirSyncTime | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last DirSync time..."
		$TextboxResults.Text = Get-MsolCompanyInformation -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List LastDirSyncTime | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getLastPasswordSyncTimeToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password sync time..."
		$TextboxResults.Text = Get-MsolCompanyInformation | Format-List LastPasswordSyncTime | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting last password sync time..."
		$TextboxResults.Text = Get-MsolCompanyInformation -TenantId $PartnerComboBox.SelectedItem.TenantID  | Format-List LastPasswordSyncTime | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getAllCompanyInfoToolStripMenuItem_Click = {
	#What to do if connected to main o365 account
	If (Get-PSSession -name mainaccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all company info..."
		$TextboxResults.Text = Get-MsolCompanyInformation | Format-List | Out-String
	}
	#What to do if connected to partner account
	ElseIf (Get-PSSession -name partneraccount -ErrorAction SilentlyContinue)
	{
		$TextboxResults.Text = "Getting all company info..."
		$TextboxResults.Text = Get-MsolCompanyInformation -TenantId $PartnerComboBox.SelectedItem.TenantID | Format-List | Out-String
	}
	Else
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Sharing Policy

$getSharingPolicyToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting all sharing policies..."
		$TextboxResults.Text = Get-SharingPolicy | Format-List Name, Domains, Enabled, Default, Identity, WhenChanged, WhenCreated, IsValid, ObjectState | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$disableASharingPolicyToolStripMenuItem_Click = {
	$DisableSharingPolicy = Read-Host "Enter the name of the policy you want to disable"
	try
	{
		$TextboxResults.Text = "Disabling the sharing policy $DisableSharingPolicy..."
		Set-SharingPolicy -Identity $DisableSharingPolicy -Enabled $false
		$TextboxResults.Text = Get-SharingPolicy -Identity $DisableSharingPolicy | Format-List Name, Enabled, ObjectState | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$enableASharingPolicyToolStripMenuItem_Click = {
	$EnableSharingPolicy = Read-Host "Enter the name of the policy you want to Enable"
	try
	{
		$TextboxResults.Text = "Enabling the sharing policy $EnableSharingPolicy..."
		Set-SharingPolicy -Identity $EnableSharingPolicy -Enabled $True
		$TextboxResults.Text = Get-SharingPolicy -Identity $EnableSharingPolicy | Format-List Name, Enabled, ObjectState | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
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
		New-SharingPolicy -Name $NewSharingPolicyName -Domains $SharingPolicy
		$TextboxResults.Text = Get-SharingPolicy -Identity $NewSharingPolicyName | Format-List Name, Domains, Enabled, Default, Identity, WhenChanged, WhenCreated, IsValid, ObjectState | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getInfoForASingleSharingPolicyToolStripMenuItem_Click = {
	$DetailedInfoForSharingPolicy = Read-Host "Enter the name of the policy you want info for"
	try
	{
		$TextboxResults.Text = "Getting info for the sharing policy $DetailedInfoForSharingPolicy..."
		$TextboxResults.Text = Get-SharingPolicy -Identity $DetailedInfoForSharingPolicy | Format-List Name, Domains, Enabled, Default, Identity, WhenChanged, WhenCreated, IsValid, ObjectState | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Configuration 

$enableCustomizationToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Enabling customization..."
		Enable-OrganizationCustomization
		$TextboxResults.Text = "Success"
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getCustomizationStatusToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting customization status..."
		$TextboxResults.Text = Get-OrganizationConfig | Format-List  Identity, IsDehydrated | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getOrganizationCustomizationToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Getting customization information..."
		$TextboxResults.Text = Get-OrganizationConfig | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getSharepointSiteToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting sharepoint URL"
		$TextboxResults.Text = Get-OrganizationConfig | Format-List SharePointUrl | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}



###Reporting###

$getAllMailboxSizesToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Generating mailbox sizes report..."
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
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Mail Malware Report

$getMailMalwareReportToolStripMenuItem1_Click = {
	$TextboxResults.Text = "Generating recent mail malware report..."
	Try
	{
		$TextboxResults.Text = Get-MailDetailMalwareReport | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getMailMalwareReportFromSenderToolStripMenuItem_Click = {
	$MalwareSender = Read-Host "Enter the email of the sender"
	try
	{
		$TextboxResults.Text = "Generating mail malware report sent from $MalwareSender..."
		$TextboxResults.Text = Get-MailDetailMalwareReport -SenderAddress $MalwareSender | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getMailMalwareReportBetweenDatesToolStripMenuItem_Click = {
	$MalwareReportStart = Read-Host "Enter the start date. EX: 06/13/2015"
	$MalwareReportEnd = Read-Host "Enter the end date. EX: 06/15/2015 "
	try
	{
		$TextboxResults.Text = "Generating mail malware report between $MalwareReportStart and $MalwareReportEnd..."
		$TextboxResults.Text = Get-MailDetailMalwareReport -StartDate $MalwareReportStart -EndDate $MalwareReportEnd | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getMailMalwareReportToARecipientToolStripMenuItem_Click = {
	$MalwareRecipient = Read-Host "Enter the email of the recipient"
	try
	{
		$TextboxResults.Text = "Generating mail malware report sent to $MalwareRecipient..."
		$TextboxResults.Text = Get-MailDetailMalwareReport -RecipientAddress $MalwareRecipient | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getMailMalwareReporforInboundToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Generating mail malware inbound report..."
		$TextboxResults.Text = Get-MailDetailMalwareReport -Direction Inbound | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getMailMalwareReportForOutboundToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = "Generating mail malware outbound report..."
		$TextboxResults.Text = Get-MailDetailMalwareReport -Direction Outbound | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Mail Traffic Report

$getRecentMailTrafficReportToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Generating recent mail traffic report..."
		$TextboxResults.Text = Get-MailTrafficReport | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getInboundMailTrafficReportToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Generating inbound traffic report..."
		$TextboxResults.Text = Get-MailTrafficReport -Direction Inbound | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getOutboundMailTrafficReportToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Generating outbound mail traffic report..."
		$TextboxResults.Text = Get-MailTrafficReport -Direction Outbound | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getMailTrafficReportBetweenDatesToolStripMenuItem_Click = {
	$MailTrafficStart = Read-Host "Enter the start date. EX: 06/13/2015"
	$MailTrafficEnd = Read-Host "Enter the end date. EX: 06/15/2015 "
	Try
	{
		$TextboxResults.Text = "Generating mail traffic report between $MailTrafficStart and $MailTrafficEnd..."
		$TextboxResults.Text = Get-MailTrafficReport -StartDate $MailTrafficStart -EndDate $MailTrafficEnd | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}



###SHARED MAILBOXES###

$createASharedMailboxToolStripMenuItem_Click = {
	$NewSharedMailbox = Read-Host "Enter the name of the new Shared Mailbox"
	Try
	{
		$TextboxResults.Text = "Creating new shared mailbox $NewSharedMailbox"
		New-Mailbox -Name $NewSharedMailbox –Shared
		$TextboxResults.Text = Get-Mailbox -RecipientTypeDetails SharedMailbox | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getAllSharedMailboxesToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting list of shared mailboxes..."
		$TextboxResults.Text = Get-Mailbox -RecipientTypeDetails SharedMailbox | Format-Table -AutoSize | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$convertRegularMailboxToSharedToolStripMenuItem_Click = {
	$ConvertMailboxtoShared = Read-Host "Enter the name of the account you want to convert"
	Try
	{
		$TextboxResults.Text = "Converting $ConvertMailboxtoShared to a shared mailbox..."
		Set-Mailbox $ConvertMailboxtoShared –Type shared
		$TextboxResults.Text = Get-Mailbox -Identity $ConvertMailboxtoShared | Format-List UserPrincipalName, DisplayName, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses, IsDirSynced | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$convertSharedMailboxToRegularToolStripMenuItem_Click = {
	$ConvertMailboxtoRegular = Read-Host "Enter the name of the account you want to convert"
	Try
	{
		$TextboxResults.Text = "Converting $ConvertMailboxtoRegular to a regular mailbox..."
		Set-Mailbox $ConvertMailboxtoRegular –Type Regular
		$TextboxResults.Text = Get-Mailbox -Identity $ConvertMailboxtoRegular | Format-List UserPrincipalName, DisplayName, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses, IsDirSynced | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getDetailedInfoForASharedMailboxToolStripMenuItem_Click = {
	$SharedMailboxDetailedInfo = Read-Host "Enter the name of the shared mailbox"
	Try
	{
		$TextboxResults.Text = "Getting shared mailbox information for $SharedMailboxDetailedInfo..."
		$TextboxResults.Text = Get-Mailbox $SharedMailboxDetailedInfo | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

	#Permissions

$addFullAccessPermissionsToASharedMailboxToolStripMenuItem_Click = {
	$SharedMailboxFullAccess = Read-Host "Enter the name of the shared mailbox"
	$GrantFullAccesstoSharedMailbox = Read-Host "Enter the UPN of the user that will have full access"
	Try
	{
		$TextboxResults.Text = "Granting Full Access permissions to $GrantFullAccesstoSharedMailbox for the $SharedMailboxFullAccess shared mailbox..."
		$TextboxResults.Text = Add-MailboxPermission $SharedMailboxFullAccess -User $GrantFullAccesstoSharedMailbox -AccessRights FullAccess -InheritanceType All | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getSharedMailboxPermissionsToolStripMenuItem_Click = {
	$SharedMailboxPermissionsList = Read-Host "Enter the name of the Shared Mailbox"
	Try
	{
		$TextboxResults.Text = "Getting Send As permissions for $SharedMailboxPermissionsList..."
		#$TextboxResults.Text = Get-RecipientPermission $SharedMailboxPermissionsList | Format-List | Out-String
		$TextboxResults.Text = Get-RecipientPermission $SharedMailboxPermissionsList | Where-Object { ($_.Trustee -notlike "NT AUTHORITY\SELF") } | Format-List Trustee, AccessControlType, AccessRights | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getSharedMailboxFullAccessPermissionsToolStripMenuItem_Click = {
	$SharedMailboxFullAccessPermissionsList = Read-Host "Enter the name of the Shared Mailbox"
	Try
	{
		$TextboxResults.Text = "Getting Full Access permissions for $SharedMailboxFullAccessPermissionsList..."
		$TextboxResults.Text = Get-MailboxPermission $SharedMailboxFullAccessPermissionsList | Where-Object { ($_.User -notlike "NT AUTHORITY\SELF") } | Format-List Identity, User, AccessRights | Out-String

	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$addSendAsAccessToASharedMailboxToolStripMenuItem_Click = {
	$SharedMailboxSendAsAccess = Read-Host "Enter the name of the shared mailbox"
	$SharedMailboxSendAsUser = Read-Host "Enter the UPN of the user"
	Try
	{
		$TextboxResults.Text = "Getting Send As permissions for $SharedMailboxSendAsAccess..."
		$TextboxResults.Text = Add-RecipientPermission $SharedMailboxSendAsAccess -Trustee $SharedMailboxSendAsUser -AccessRights SendAs | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
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
		New-MailContact -Name $ContactName -FirstName $ContactFirstName -LastName $ContactsLastName -ExternalEmailAddress $ContactExternalEmail
		$TextboxResults.Text = Get-MailContact -Identity $ContactName | Format-List DisplayName, EmailAddresses, PrimarySmtpAddress, ExternalEmailAddress, RecipientType | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getAllContactsToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all contacts..."
		$TextboxResults.Text = Get-MailContact | Format-List DisplayName, EmailAddresses, PrimarySmtpAddress, ExternalEmailAddress, RecipientType | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$getDetailedInfoForAContactToolStripMenuItem_Click = {
	$DetailedInfoForContact = Read-Host "Enter the contact name, displayname, alias or email address"
	Try
	{
		$TextboxResults.Text = "Getting detailed info for $DetailedInfoForContact..."
		$TextboxResults.Text = Get-MailContact -Identity $DetailedInfoForContact | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$removeAContactToolStripMenuItem_Click = {
	$RemoveMailContact = Read-Host "Enter the contact name, displayname, alias or email address"
	Try
	{
		$TextboxResults.Text = "Removing contact $RemoveMailContact..."
		Remove-MailContact -Identity $RemoveMailContact
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

	#Global Address List

$hideContactFromGALToolStripMenuItem_Click = {
	$HideGALMailContact = Read-Host "Enter the contact name, displayname, alias or email address"
	Try
	{
		$TextboxResults.Text = "Hiding $HideGALMailContact from the GAL..."
		Set-MailContact -Identity $HideGALMailContact -HiddenFromAddressListsEnabled $true
		$TextboxResults.Text = Get-MailContact -Identity $HideGALMailContact | Format-List DisplayName, HiddenFromAddressListsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
}

$unhideContactFromGALToolStripMenuItem_Click = {
	$unHideGALMailContact = Read-Host "Enter the contact name, displayname, alias or email address"
	Try
	{
		$TextboxResults.Text = "unhiding $unHideGALMailContact from the GAL..."
		Set-MailContact -Identity $unHideGALMailContact -HiddenFromAddressListsEnabled $False
		$TextboxResults.Text = Get-MailContact -Identity $unHideGALMailContact | Format-List DisplayName, HiddenFromAddressListsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getGALStatusForAllUsersToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting GAL status for all users..."
		$TextboxResults.Text = Get-MailContact | Format-List DisplayName, HiddenFromAddressListsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getContactsHiddenFromGALToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all users that are hidden from the GAL..."
		$TextboxResults.Text = Get-MailContact | Where-Object { $_.HiddenFromAddressListsEnabled -like "True" } | Format-List DisplayName, HiddenFromAddressListsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}

$getContactsNotHiddenFromGALToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.Text = "Getting all users that not are hidden from the GAL"
		$TextboxResults.Text = Get-MailContact | Where-Object { $_.HiddenFromAddressListsEnabled -like "False" } | Format-List DisplayName, HiddenFromAddressListsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		$TextboxResults.Text = ""
	}
	
}



###HELP###

	#About

$aboutToolStripMenuItem_Click = {
	$TextboxResults.Text = "                 o365 Administration Center v2.0.1 
	
	HOW TO USE
To start, click the Connect to Office 365 button. This will connect you 
to Exchange Online using Remote PowerShell. Once you are connected the 
button will grey out and the form title will change to -CONNECTED TO O365-

The TextBox will display all output for each command. If nothing appears
and there was no error then the result was null. The Textbox also serves
as input, passing your own commands to PowerShell with the result 
populating in the same Textbox. To run your own command simply clear the
Textbox and enter in your command and press the Run Command button or
press Enter on your keyboard.

You can also export the results to a file using the Export to File 
button. The Textbox also allows copy and paste. The Exit button will
properly end the Remote PowerShell session"
	
}

	#Pre-Reqs

$prerequisitesToolStripMenuItem_Click = {
	$TextboxResults.Text = "
Windows Versions:
Windows 7 Service Pack 1 (SP1) or higher
Windows Server 2008 R2 SP1 or higher

You need to install the Microsoft.NET Framework 4.5 or later and 
then Windows Management Framework 3.0 or later. 

MS-Online module needs to be installed. Install the MSOnline Services
Sign In Assistant: 
https://www.microsoft.com/en-us/download/details.aspx?id=41950 

Azure Active Directory Module for Windows PowerShell needs to be 
installed: http://go.microsoft.com/fwlink/p/?linkid=236297

The Microsoft Online Services Sign-In Assistant provides end user
sign-in capabilities to Microsoft Online Services, such as Office 365.

Windows PowerShell needs to be configured to run scripts, and by default, 
it isn't. To enable Windows PowerShell to run scripts, run the following 
command in an elevated Windows PowerShell window (a Windows PowerShell
window you open by selecting Run as administrator): 
Set-ExecutionPolicy Unrestricted"

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
