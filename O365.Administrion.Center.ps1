<#
.SYNOPSIS
This is the source code for o365 Adminsitration Center
.DESCRIPTION
The o365 Admin Center is a GUI application that administrators can use to perform some of the most common o365 tasks. The output (error or success) is sent to the textbox which also acts as a input for custom commands. You can also save the output to a file. 
.NOTES
This is built with a GUI and not a stand alone script.
.LINK
www.bwya77.com
#>

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
}

$ButtonExit_Click= {
	#Disconnects O365 Session
	Get-PSSession | Remove-PSSession
	
	<# Creates a pop up box telling the user they are disconnected from the o365 session. This is commented out as it will show True every time as the command will never error out even if there 
	is no session to disconnect from #>
	#[void][System.Windows.Forms.MessageBox]::Show("You are disconnected from O365", "Message")
}

$ButtonConnectTo365_Click = {
	try
	{
		$o365creds = (Get-Credential -Message "O365 Credentials")
		#CONNECT TO OFFICE365
		Connect-MsolService -Credential $o365creds
		
		#CONNECT TO EXCHANGE ONLINE
		$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $o365creds -Authentication Basic -AllowRedirection
		Import-PSSession $exchangeSession -DisableNameChecking
		#Disable Button
		$ButtonConnectTo365.Enabled = $false
		#Sets custom button text
		$ButtonConnectTo365.Text = "Connected to O365"
		#Sets custom form text
		$FormO365AdministrationCenter.Text = "-Connected to O365-"
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("Error: You are not connected to O365, Please verify the correct username and password", "Error")
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
				$TextboxResults.text = Invoke-Expression $userinput | Format-List | Out-String
			}
			Catch
			{
				[System.Windows.Forms.MessageBox]::Show("$_", "Error")
			}
}


###QUOTA MENU ITEMS###

$getUserQuotaToolStripMenuItem_Click={
	$QuotaUser = Read-Host "Enter the Email of the user you want to view Quota information for"
	try
	{
		$TextboxResults.text = Get-Mailbox $QuotaUser | Fl *Quota | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}


###LICENSED MENU ITEMS###

$getLicensedUsersToolStripMenuItem_Click={
	try
	{
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.isLicensed -eq "TRUE" } | Format-List DisplayName, Licenses | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$displayAllUsersWithoutALicenseToolStripMenuItem_Click = {
	Try
	{
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.isLicensed -like "False" } | Format-List UserPrincipalName | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$removeAllUnlicensedUsersToolStripMenuItem_Click = {
	Try
	{
		Get-MsolUser -all | Where-Object { $_.isLicensed -ne "true" } | Remove-MsolUser -Force
		$TextboxResults.text = Get-MsolUser | Where-Object { $_.isLicensed -like "False" } | Format-List UserPrincipalName | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}


###CALENDAR MENU ITEMS###

$addCalendarPermissionsToolStripMenuItem_Click={
	$Calendaruser = Read-Host "Calendar you want to give access to"
	$Calendaruser2 = Read-Host "User to give access to"
	$TextboxResults.text = "Calendar Permissions: Owner; PublishingEditor; PublishingAuthor; Reviewer"
	$level = Read-Host "Access Level?"
	try
	{
		$TextboxResults.text = Set-MailboxFolderPermission -Identity ${Calendaruser}:\calendar -user $Calendaruser2 -AccessRights $level | Format-List Identity, FolderName, User, AccessRights, IsValid | Out-String -ErrorAction Stop
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$viewUsersCalendarPermissionsToolStripMenuItem_Click = {
	$CalUserPermissions = Read-Host "What user would you like calendar permissions for?"
	Try
	{
		$TextboxResults.Text = Get-MailboxFolderPermission -Identity ${CalUserPermissions}:\calendar | Format-List Identity, FolderName, User, AccessRights, IsValid, ObjectSpace | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$addAllUsersPermissionToASingleCalrndarToolStripMenuItem_Click = {
	$users = Get-Mailbox | Select -ExpandProperty Alias
	$AllCalUser = Read-Host "Which user's calendar would you like everyone to have access to? Please enter the full email address"
	$TextboxResults.text = "Calendar Permissions: Owner; PublishingEditor; PublishingAuthor; Reviewer"
	$level2 = Read-Host "Access Level?"
	try
	{
		$TextboxResults.Text = Foreach ($user in $users) { Add-MailboxFolderPermission ${AllCalUser}:\calendar -user $user -accessrights $level2 }﻿ | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}


###CLUTTER MENU ITEMS###

$disableClutterForAllToolStripMenuItem_Click={
	try
	{
		$TextboxResults.text = Get-Mailbox | Set-Clutter -Enable $false | Format-List MailboxIdentity, IsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$enableClutterForAllToolStripMenuItem_Click={
	try
	{
		$TextboxResults.text = Get-Mailbox | Set-Clutter -Enable $True | Format-List MailboxIdentity, IsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$enableClutterForAUserToolStripMenuItem_Click = {
	$UserEnableClutter = Read-Host "Which user would you like to enable Clutter for?"
	try
	{
		$TextboxResults.text = Get-Mailbox $UserEnableClutter | Set-Clutter -Enable $True | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$disableClutterForAUserToolStripMenuItem_Click = {
	$UserDisableClutter = Read-Host "Which user would you like to disable Clutter for?"
	try
	{
		$TextboxResults.text = Get-Mailbox $UserDisableClutter | Set-Clutter -Enable $False | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$getClutterInfoForAUserToolStripMenuItem_Click = {
	$GetCluterInfoUser = Read-Host "What user would you like to view Clutter information about?"
	Try
	{
		$TextboxResults.Text = Get-Clutter -Identity $GetCluterInfoUser | Format-List IsEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}


#DISTRIBUTION GROUP MENU ITEMS###

$displayDistributionGroupsToolStripMenuItem_Click={
	try
	{
		$TextboxResults.text = Get-DistributionGroup | Format-List DisplayName, SamAccountName, GroupType, IsDirSynced, EmailAddresses | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$createADistributionGroupToolStripMenuItem_Click = {
	$NewDistroGroup = Read-Host "What is the name of the new distribution group?"
	try
	{
		$TextboxResults.Text = New-DistributionGroup -Name $NewDistroGroup | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$deleteADistributionGroupToolStripMenuItem_Click = {
	$DeleteDistroGroup = Read-Host "Enter the name of the Distribtuion group you want deleted."
	try
	{
		Remove-DistributionGroup $DeleteDistroGroup
		$TextboxResults.text = Get-DistributionGroup | Format-List DisplayName | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$allowDistributionGroupToReceiveExternalEmailToolStripMenuItem_Click = {
	#untested code
	$AllowExternalEmail = Read-Host "Enter the name of the Distribtuion Group you want to allow external email to"
	try
	{
		Set-DistributionGroup $AllowExternalEmail -RequireSenderAuthenticationEnabled $False 
		$TextboxResults.text = Get-DistributionGroup $AllowExternalEmail | Format-List Name, RequireSenderAuthenticationEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$hideDistributionGroupFromGALToolStripMenuItem_Click = {
	$GroupHideGAL = Read-Host "Enter the name of the Distribtuion Group you want to allow external email to"
	try
	{
		Set-DistributionGroup $GroupHideGAL -HiddenFromAddressListsEnabled $True
		$TextboxResults.text = Get-DistributionGroup $GroupHideGAL | Format-List Name, HiddenFromAddressListsEnabled | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$displayDistributionGroupMembersToolStripMenuItem_Click = {
	$ListDistributionGroupMembers = Read-Host "Enter the name of the Distribution Group you want to list members of"
	try
	{
		$TextboxResults.Text = Get-DistributionGroupMember $ListDistributionGroupMembers | Format-List Name | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}


###USERS GENERAL MENU ITEMS###

$getListOfUsersToolStripMenuItem_Click={
	$TextboxResults.text = Get-MSOLUser | Format-List UserPrincipalName | Out-String
	
}

$getDetailedInfoForAUserToolStripMenuItem_Click = {
	$DetailedInfoUser = Read-Host "Enter the User Principal Name of the user you want more information about"
	try
	{
		$TextboxResults.text = Get-MsolUser -UserPrincipalName $DetailedInfoUser | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$creatOutOfOfficeAutoReplyForAUserToolStripMenuItem_Click = {
	$OOOautoreplyUser = Read-Host "What user is the Out Of Office auto reply for?"
	$OOOInternal = Read-Host "What is the Internal Message"
	$OOOExternal = Read-Host "What is the External Message"
	Try
	{
		Set-MailboxAutoReplyConfiguration -Identity $OOOautoreplyUser -AutoReplyState Enabled -ExternalMessage $OOOExternal -InternalMessage $OOOInternal
		$TextboxResults.Text = Get-MailboxAutoReplyConfiguration -Identity $OOOautoreplyUser | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$changeUsersLoginNameToolStripMenuItem_Click = {
	$UserChangeUPN = Read-Host "What user would you like to change their login name for? Enter their UPN"
	$NewUserUPN = Read-Host "What would you like the new username to be?"
	Try
	{
		Set-MsolUserPrincipalname -UserPrincipalName $UserChangeUPN -NewUserPrincipalName $NewUserUPN
		$TextboxResults.text = Get-MSOLUser -UserPrincipalName $NewUserUPN | Format-List UserPrincipalName | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$deleteAUserToolStripMenuItem_Click = {
	$DeleteUser = Read-Host "Enter the UPN of the user you want to delete"
	Try
	{
		$TextboxResults.text = Remove-MsolUser –UserPrincipalName $DeleteUser | Format-List UserPrincipalName | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$createANewUserToolStripMenuItem_Click = {
	$Firstname = Read-Host "Enter the First Name for the new user"
	$LastName = Read-Host "Enter the Last Name for the new user"
	$DisplayName = Read-Host "Enter the Display Name for the new user"
	$NewUser = Read-Host "Enter the UPN for the new user"
	Try
	{
		$TextboxResults.text = New-MsolUser -UserPrincipalName $NewUser -FirstName $Firstname -LastName $LastName -DisplayName $DisplayName | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}


###PASSWORD MENU ITEMS###

$enableStrongPasswordForAUserToolStripMenuItem_Click={
	$UserEnableStrongPasswords = Read-Host "Enter the User Principal Name of the user you want to enable strong password policy for"
		try
		{
			Set-MsolUser -UserPrincipalName $UserEnableStrongPasswords -StrongPasswordRequired $True
			$TextboxResults.text = Get-MsolUser -UserPrincipalName $UserEnableStrongPasswords | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
		}
		Catch
		{
			[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
	
}

$getAllUsersStrongPasswordPolicyInfoToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.text = Get-MsolUser | Format-List userprincipalname, strongpasswordrequired | Out-String
		
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$disableStrongPasswordsForAUserToolStripMenuItem_Click={
	$UserdisableStrongPasswords = Read-Host "Enter the User Principal Name of the user you want to disable strong password policy for"
		try
		{
			Set-MsolUser -UserPrincipalName $UserdisableStrongPasswords -StrongPasswordRequired $False
			$TextboxResults.text = Get-MsolUser -UserPrincipalName $UserdisableStrongPasswords | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
		}
		Catch
		{
			[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
	
}

$enableStrongPasswordsForAllToolStripMenuItem_Click = {
	try
	{
		Get-MsolUser | Set-MsolUser -StrongPasswordRequired $True
		$TextboxResults.text = Get-MsolUser | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$disableStrongPasswordsForAllToolStripMenuItem_Click = {
	try
	{
		Get-MsolUser | Set-MsolUser -StrongPasswordRequired $False
		$TextboxResults.text = Get-MsolUser | Format-List UserPrincipalName, StrongPasswordRequired | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$resetPasswordForAUserToolStripMenuItem_Click = {
	$ResetPasswordUser = Read-Host "What user would you like to reset the password for?"
	$NewPassword = Read-Host "What would you like the new password to be?"
	try
	{
		Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser –NewPassword $NewPassword -ForceChangePassword $False
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$setPasswordToNeverExpireForAllToolStripMenuItem_Click = {
	try
	{
		Get-MsolUser | Set-MsolUser –PasswordNeverExpires $True
		$TextboxResults.text = Get-MSOLUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$setPasswordToExpireForAllToolStripMenuItem_Click = {
	try
	{
		Get-MsolUser | Set-MsolUser –PasswordNeverExpires $False | Format-List | Out-String
		$TextboxResults.text = Get-MSOLUser | Format-List UserPrincipalName, PasswordNeverExpires | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$resetPasswordForAllToolStripMenuItem_Click = {
	$SetPasswordforAll = Read-Host "What password would you like to set for all users?"
	try
	{
		Get-MsolUser | %{ Set-MsolUserPassword -userPrincipalName $_.UserPrincipalName –NewPassword $SetPasswordforAll -ForceChangePassword $False }
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$setATempPasswordForAllToolStripMenuItem_Click = {
	$SetTempPasswordforAll = Read-Host "What password would you like to set for all users?"
	try
	{
		Get-MsolUser | Set-MsolUserPassword –NewPassword $SetTempPasswordforAll -ForceChangePassword $True
		$TextboxResults.Text = "Temporary password has been set to $SetTempPasswordforAll Please note that users will be prompted to change it upon first logon"
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}


###MAILBOX PERMISSIONS MENU ITEMS###

$addFullPermissionsToAMailboxToolStripMenuItem_Click = {
	$mailboxAccess = read-host "Mailbox you want to give full-access to"
	$mailboxUser = read-host "Enter the email of the user that will have full access"
		try
		{
			$TextboxResults.text = Add-MailboxPermission $mailboxAccess -User $mailboxUser -AccessRights FullAccess -InheritanceType All | Format-List | Out-String
		}
		Catch
		{
			[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
}

$addSendAsPermissionToAMailboxToolStripMenuItem_Click={
	$SendAsAccess = read-host "Mailbox you want to give Send As access to"
	$mailboxUserAccess = read-host "Enter the user that will have Send As access"
		try
		{
			$TextboxResults.text = Add-RecipientPermission $SendAsAccess -Trustee $mailboxUserAccess -AccessRights SendAs -Confirm:$False | Format-List | Out-String
		}
		Catch
		{
			[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
	
}

$assignSendOnBehalfPermissionsForAMailboxToolStripMenuItem_Click = {
	$SendonBehalfof = read-host "Mailbox you want to give Send As access to"
	$mailboxUserSendonBehalfAccess = read-host "Enter the user that will have Send As access"
		try
		{
			Set-Mailbox -Identity $SendonBehalfof -GrantSendOnBehalfTo $mailboxUserSendonBehalfAccess
			$TextboxResults.text = Get-Mailbox -Identity $SendonBehalfof | Format-List Identity, GrantSendOnBehalfTo | Out-String
		}
		Catch
		{
			[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
}

$displayMailboxPermissionsForAUserToolStripMenuItem_Click={
	$MailboxUserFullAccessPermission = Read-Host "Which user would you like to view Full Access permissions for?"
		try
		{
			$TextboxResults.text = Get-MailboxPermission $MailboxUserFullAccessPermission | Where { ($_.IsInherited -eq $False) -and -not ($_.User -like “NT AUTHORITY\SELF”) } | Format-List AccessRights, Deny, InheritanceType, User, Identity, IsInherited, IsValid | Out-String
		}
		Catch
		{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
}

$displaySendAsPermissionForAMailboxToolStripMenuItem_Click={
	$MailboxUserSendAsPermission = Read-Host "Which user would you like to view Send As permissions for?"
		try
		{
			$TextboxResults.text = Get-RecipientPermission $MailboxUserSendAsPermission | Where { ($_.IsInherited -eq $False) -and -not ($_.Trustee -like “NT AUTHORITY\SELF”) } | Format-List | Out-String
		}
		Catch
		{
			[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
	
}

$displaySendOnBehalfPermissionsForMailboxToolStripMenuItem_Click={
	$MailboxUserSendonPermission = Read-Host "Which user would you like to view Send On Behalf Of permission for?"
		try
		{
			$TextboxResults.text = Get-RecipientPermission $MailboxUserSendonPermission | Where { ($_.IsInherited -eq $False) -and -not ($_.Trustee -like “NT AUTHORITY\SELF”) } | Format-List | Out-String
		}
		Catch
		{
			[System.Windows.Forms.MessageBox]::Show("$_", "Error")
		}
	
}

$removeFullAccessPermissionsForAMailboxToolStripMenuItem_Click={
	$UserRemoveFullAccessRights = Read-Host "What user mailbox would you like modify Full Access rights to"
	$RemoveFullAccessRightsUser = Read-Host "Which user would you like to remove"
	try
	{
		Remove-MailboxPermission  $UserRemoveFullAccessRights -User $RemoveFullAccessRightsUser -AccessRights FullAccess -Confirm:$False -ea 1
		$TextboxResults.text = Get-MailboxPermission $UserRemoveFullAccessRights | Where { ($_.IsInherited -eq $False) -and -not ($_.User -like “NT AUTHORITY\SELF”) } | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$revokeSendAsPermissionsForAMailboxToolStripMenuItem_Click={
	$UserDeleteSendAsAccessOn = Read-Host "What user mailbox would you like to modify Send As permission for?"
	$UserDeleteSendAsAccess = Read-Host "What user would you like to remove from having Send As access to?"
	try
	{
		$TextboxResults.Text = Remove-RecipientPermission $UserDeleteSendAsAccessOn -AccessRights SendAs -Trustee $UserDeleteSendAsAccess -Confirm:$False | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}


### RESOURCE MAILBOX###

$convertAMailboxToARoomMailboxToolStripMenuItem_Click = {
	$MailboxtoRoom = Read-Host "What user would you like to convert to a Room Mailbox? Please enter the full email address"
	Try
	{
		Set-Mailbox $MailboxtoRoom -Type Room
		$TextboxResults.Text = Get-MailBox $MailboxtoRoom | Format-List Name, ResourceType, PrimarySmtpAddress, EmailAddresses, UserPrincipalName, IsMailboxEnabled | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$enableAutomaticBookingForAllResourceMailboxToolStripMenuItem1_Click = {
	Try
	{
		Get-MailBox | Where { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AutomateProcessing:AutoAccept
		$TextboxResults.Text = Get-MailBox | Where { $_.ResourceType -eq "Room" } | Get-CalendarProcessing | Format-List Identity, AutomateProcessing | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$denyConflictMeetingsForAllResourceMailboxesToolStripMenuItem_Click = {
	Try
	{
		Get-MailBox | Where { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AllowConflicts $False
		$TextboxResults.Text = Get-MailBox | Where { $_.ResourceType -eq "Room" } | Get-CalendarProcessing | Format-List Identity, AllowConflicts | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$createANewRoomMailboxToolStripMenuItem_Click = {
	$TextboxResults.Text = "Description: Create a new Room mailbox"
	$NewRoomMailbox = Read-Host "Enter the name of the new room mailbox"
	Try
	{
		$TextboxResults.Text = New-Mailbox -Name $NewRoomMailbox -Room | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$disallowconflictmeetingsToolStripMenuItem_Click = {
	$ConflictMeetingDeny = Read-Host "Enter the Room Name of the Resource Calendar you want to disallow conflicts"
	try
	{
		Set-CalendarProcessing $ConflictMeetingDeny -AllowConflicts $False
		$TextboxResults.Text = Get-CalendarProcessing -identity $ConflictMeetingDeny | Format-List Identity, AllowConflicts | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$allowConflictMeetingsToolStripMenuItem_Click = {
	$ConflictMeetingAllow = Read-Host "Enter the Room Name of the Resource Calendar you want to allow conflicts"
	try
	{
		Set-CalendarProcessing $ConflictMeetingAllow -AllowConflicts $True
		$TextboxResults.Text = Get-CalendarProcessing -identity $ConflictMeetingAllow | Format-List Identity, AllowConflicts | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$getListOfRoomMailboxesToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = $TextboxResults.Text = Get-MailBox | Where { $_.ResourceType -eq "Room" } | Format-List Identity, PrimarySmtpAddress, EmailAddresses, UserPrincipalName | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}


###MAILBOX PERMISSIONS MENU ITEMS###

$displayAllDeletedUsersToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$deleteAllUsersInRecycleBinToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin –Force | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$deleteSpecificUsersInRecycleBinToolStripMenuItem_Click = {
	$DeletedUserRecycleBin = Read-Host "Please enter the User Principal Name of the user you want to permanently delete"
	try
	{
		Remove-MsolUser -UserPrincipalName $DeletedUserRecycleBin -RemoveFromRecycleBin -Force
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$restoreDeletedUserToolStripMenuItem_Click = {
	$RestoredUserFromRecycleBin = Read-Host "Enter the User Principal Name of the user you want to restore"
	try
	{
		Restore-MsolUser –UserPrincipalName $RestoredUserFromRecycleBin -AutoReconcileProxyConflicts
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}


###SPAM###

$checkSafeAndBlockedSendersForAUserToolStripMenuItem_Click = {
	$CheckSpamUser = Read-Host "Enter the UPN of the user you want to check blocked and allowed senders for"
	try
	{
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $CheckSpamUser | Format-List Identity, TrustedListsOnly, ContactsTrusted, TrustedSendersAndDomains, BlockedSendersAndDomains, TrustedRecipientsAndDomains, IsValid | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$blacklistDomainForAllToolStripMenuItem_Click = {
	$BlacklistDomain = Read-Host "Enter the domain you want to blacklist for all users"
	try
	{
		$TextboxResults.Text = Get-Mailbox | Set-MailboxJunkEmailConfiguration -BlockedSendersAndDomains $BlacklistDomain | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$whitelistDomainForAllToolStripMenuItem_Click = {
	$AllowedDomain = Read-Host "Enter the domain you want to whitelist for all users"
	try
	{
		$TextboxResults.Text = Get-Mailbox | Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains $AllowedDomain | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$whitelistDomainForASingleUserToolStripMenuItem_Click = {
	$Alloweddomainuser = Read-Host "Enter the UPN of the user you want to modify junk email for"
	$AllowedDomain2 = Read-Host "Enter the domain you want to whitelist"
	try
	{
		Set-MailboxJunkEmailConfiguration -Identity $Alloweddomainuser -TrustedSendersAndDomains $AllowedDomain2
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $Alloweddomainuser | Format-List Identity, TrustedSendersAndDomains, TrustedRecipientsAndDomains | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	} 
}

$blacklistDomainForASingleUserToolStripMenuItem_Click = {
	$Blockeddomainuser = Read-Host "Enter the UPN of the user you want to modify junk email for"
	$BlockedDomain2 = Read-Host "Enter the domain you want to blacklist"
	try
	{
		Set-MailboxJunkEmailConfiguration -Identity $Blockeddomainuser -BlockedSendersAndDomains $BlockedDomain2
		$TextboxResults.Text = Get-MailboxJunkEmailConfiguration -Identity $Blockeddomainuser | Format-List Identity, BlockedSendersAndDomains | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}


###JUNK ITEMS###
$menustrip1_ItemClicked=[System.Windows.Forms.ToolStripItemClickedEventHandler]{
#Event Argument: $_ = [System.Windows.Forms.ToolStripItemClickedEventArgs]
	#TODO: Place custom script here
	
}

$TextboxResults_TextChanged = {
	#Left Blank
}

$allowedDomainsToolStripMenuItem_Click={
	#TODO: Place custom script here
	
}
