# O365 Administration Center

v2.0.0

####Description
O365 Administration Center is an application written in PowerShell that lets administrators easily and quickly manage their Office 365 environment. It allows partner accounts to connect to all of their tenants and run the same commands and then disconnect back to their partner account where they can then connect to another tenant. The results are shown in the textbox that also accepts custom commands. The commands can be typed in, and by pressing the enter key or the "Run Command" button the command is passed through to PowerShell and the results are displayed back on the same textbox. Results can also be exported to a file easily by using the "Export to File" button which uses the Out-File cmdlet. You can end you PSSession properly by pressing the "Exit" button which will run the following command: Get-PSSession | Remove-PSSession to saftely remove your session.

Included is the .exe if you just want to run it, or the .msi if you want to install it.
___

####Prerequisites

Windows Versions:

Windows 7 Service Pack 1 (SP1) or higher

Windows Server 2008 R2 SP1 or higher

You need to install the Microsoft.NET Framework 4.5 or later and then Windows Management Framework 3.0 or later. 

MS-Online module needs to be installed. Install the MSOnline Services Sign In Assistant: https://www.microsoft.com/en-us/download/details.aspx?id=41950 

Azure Active Directory Module for Windows PowerShell needs to be installed: http://go.microsoft.com/fwlink/p/?linkid=236297

The Microsoft Online Services Sign-In Assistant provides end user sign-in capabilities to Microsoft Online Services, such as Office 365.

Windows PowerShell needs to be configured to run scripts, and by default, it isn't. To enable Windows PowerShell to run scripts, run the following command in an elevated Windows PowerShell window (a Windows PowerShell window you open by selecting Run as administrator):
Set-ExecutionPolicy Unrestricted"

####O365 GUI
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Main_GUI.png)

####Login to O365
You are prompted for O365 credentials. It will then load all Exchange Online cmdlets.
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Log_In.png)

####Examples
######Get License Info
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Get-Lic_info2.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Licenses_InUse.png)

######Get Mailbox Size Report
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/MailBox_Size_Start.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Mailbox_Size_Report.png)

Results are sorted with the biggest mailboxes at the top and the smallest on at the bottom
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Mailbox_Size_Report_Results.png)


######Connect to Tenant

The default domain for each tenant will be populated in the combobox. If you do not have a partner account or have no tenants the combobox will remain empty
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Tenant_List.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Connecting_To_Partner.png)

Once you are connected, the Application Title will change to let you know what client you are managing. The Combobox will be unavailable and the "Connect to Partner" button will also be unavailable at this time.
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Connected_To_Partner.png)

######Disconnect from Tenant

To disconnect to partner and go back to managing you partner account press the "Disconnect from Partner"
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Disconnecting_Partner.png)


####Custom Commands
You can enter your own command simply by typing it into the textbox. It will pass it through to PowerShell and display the results

![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Custom_Command.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Custom_Command_Results.png)

