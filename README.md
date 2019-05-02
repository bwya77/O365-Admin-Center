# O365 Admin Center


v4.0.0

___

## .Links:

- [Website] (http://o365admin.center/)
- [Changelog] (http://o365admin.center/changelog/)
- [Community] (http://o365admin.center/community/)
- [FAQ] (http://o365admin.center/FAQ/)

## .Install instructions:
You need to install these compononents in powershell first
Set-ExecutionPolicy RemoteSigned
Install-Module MSOnline
Install-Module -Name AzureAD
___

## .MFA
MFA requires two modules:
- MSOnline 
- EXOPPSSession (downloaded from your tenant)
Once you download and install the modules, enable MFA by going to Tools> 2FA > Enable 2FA
It changes a reg key at HKEY_Current_User/Software/O365 Admin Center

## .Description
The O365 Admin Center is an application written mainly in PowerShell that lets administrators easily and quickly manage their Office 365 environment. It allows partner accounts to connect to all of their tenants and run the same commands. The results are shown in the textbox that also accepts custom commands as input. The commands can be typed in, and by pressing the enter key or the “Run Command” button the command is passed through to PowerShell and the results are displayed back on the same textbox. You can manage Exchange Online, Skype For Business, SharePoint and Compliance Center.

___

## .Screenshots

#### O365 Admin Center
![alt tag](https://www.o365admin.center/wp-content/uploads/2016/05/output_PtRs59.gif)
