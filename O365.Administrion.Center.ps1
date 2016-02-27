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

#untested
$resetPasswordForAUserToolStripMenuItem_Click = {
	$ResetPasswordUser = Write-Host "What user would you like to reset the password for?"
	$NewPassword = Write-Host "What would you like the new password to be?"
	try
	{
		$TextboxResults.text = Set-MsolUserPassword –UserPrincipalName $ResetPasswordUser –NewPassword $NewPassword -ForceChangePassword $False
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

#untested
$setPasswordToNeverExpireForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.text = Get-MsolUser | Set-MsolUser –PasswordNeverExpires $True
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

#untested
$setPasswordToExpireForAllToolStripMenuItem_Click = {
	try
	{
		$TextboxResults.text = Get-MsolUser | Set-MsolUser –PasswordNeverExpires $False
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

#untested
$resetPasswordForAllToolStripMenuItem_Click = {
	$SetPasswordforAll = Write-Host "What password would you like to set for all users?"
	try
	{
	Get-MsolUser | %{ Set-MsolUserPassword -userPrincipalName $_.UserPrincipalName –NewPassword $SetPasswordforAll -ForceChangePassword $False }
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

#untested
$setATempPasswordForAllToolStripMenuItem_Click = {
	$SetTempPasswordforAll = Write-Host "What password would you like to set for all users?"
	try
	{
		Get-MsolUser | Set-MsolUserPassword –NewPassword $SetTempPasswordforAll -ForceChangePassword $False
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
	$TextboxResults.Text = "Description: Convert a regular mailbox to a Room Mailbox"
	$MailboxtoRoom = Read-Host "What user would you like to convert to a Room Mailbox? Please enter the full email address"
	Try
	{
		$TextboxResults.Text = Set-Mailbox $MailboxtoRoom -Type Room | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$enableAutomaticBookingForAllResourceMailboxToolStripMenuItem_Click = {
	$TextboxResults.Text = "Description: Enable automatic booking for all room mailboxes. This will make it so the room mailbox will auto accept calendar invitations"
	Try
	{
		$TextboxResults.Text = Get-MailBox | Where { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AutomateProcessing:AutoAccept | Format-List | Out-String
	}
	Catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
}

$denyConflictMeetingsForAllResourceMailboxesToolStripMenuItem_Click = {
	$TextboxResults.Text = "Description: Deny conflict meetings when using the option of automatic booking"
	Try
	{
		Get-MailBox | Where { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AllowConflicts $False | Format-List | Out-String
	}
	Catch
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
		$TextboxResults.Text = Remove-MsolUser -UserPrincipalName $DeletedUserRecycleBin -RemoveFromRecycleBin -Force
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$restoreDeletedUserToolStripMenuItem_Click = {
	#Untested but known working command, not sure of the output so next line is listing the deleted items 
	$RestoredUserFromRecycleBin = Read-Host "Enter the User Principal Name of the user you want to restore"
	$NewRestoredUserPrincipalName = Read-Host "Enter the new User Principal Name of restored user"
	try
	{
		$TextboxResults.Text = Restore-MsolUser –UserPrincipalName $RestoredUserFromRecycleBin -AutoReconcileProxyConflicts -NewUserPrincipalName $NewRestoredUserPrincipalName
		$TextboxResults.Text = Get-MsolUser -ReturnDeletedUsers | Format-List UserPrincipalName, ObjectID | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
	
}

###MAILBOX PERMISSIONS MENU ITEMS###

$enableAutomaticBookingForAllResourceMailboxToolStripMenuItem1_Click = {
	#untested code
	try
	{
		$TextboxResults.Text = Get-MailBox | Where { $_.ResourceType -eq "Room" } | Set-CalendarProcessing -AutomateProcessing:AutoAccept | Format-List  | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$allowConflictMeetingsToolStripMenuItem_Click = {
	#untested code
	$ConflictMeetingAllow = Read-Host "Enter the Room Name of the Resource Calendar you want to allow conflicts"
	try
	{
		$TextboxResults.Text = Set-CalendarProcessing $ConflictMeetingAllow -AllowConflicts $True | Format-List | Out-String
	}
	catch
	{
		[System.Windows.Forms.MessageBox]::Show("$_", "Error")
	}
	
}

$disToolStripMenuItem_Click = {
	#untested code
	$ConflictMeetingDeny = Read-Host "Enter the Room Name of the Resource Calendar you want to disallow conflicts"
	try
	{
		$TextboxResults.Text = Set-CalendarProcessing $ConflictMeetingDeny -AllowConflicts $False | Format-List | Out-String
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




