# O365 Administration Center

v1.0.0

####Description
The O365 Admin Center is a GUI application that administrators can use to perform some of the most common O365 tasks. The output (error or success) is sent to the textbox which also acts as a input for custom commands. You can also save the output to a file. You can end you PSSession properly by pressing the Exit button which will run the following command: Get-PSSession | Remove-PSSession

Included is the .exe if you just want to run it, or the .msi if you want to install it.

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
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/O365_GUI2.png)

####Login to O365
You are prompted for O365 credentials. It will then load all Exchange Online cmdlets. When you sucessfully connect the form title will change and the Connect to Office 365 button will be grayed out.

![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/GUI_Logon.png)

####Examples
######Display Licensed Users
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Example_GetLicensedUsers_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Example_GetLicensedUsers_Output_.png)

######Get List of Users
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Example_GetUsers_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Example_GetUsers_Output_.png)

######Get Detailed Info for a User
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Example_GetDetailsUserInfo_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Example_GetDetailsUserInfo_Prompt_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Example_GetDetailsUserInfo_Output_.png)

####UI Elements
######Admin
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Root_Admin_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Admin_ActiveSync_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Admin_OWA_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Admin_PowerShellAccess_.png)

######Groups
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Root_Groups_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Group_DistributionGroups_.png)

######Junk Email
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Root_JunkEmail_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/JunkEmail_Blacklist_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/JunkEmail_Whitelist_.png)

######Quarantine
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Root_Quarantine_.png)

######Resource Mailbox
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Root_ResourceMailbox_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/ResourceMailbox_BookingOptions_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/ResourceMailbox_RoomMailbox2_.png)

######Users
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Root_Users_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Users_CalendarPermissions_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Users_Clutter.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Users_Licenses_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Users_MailboxPermissions_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Users_Passwords_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Users_Quota_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Users_RecycleBin.png)

######Help
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Root_Help_.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/Help_About_Output.png)

####Custom Commands
You can enter your own command simply by typing it into the textbox. It will pass it through to PowerShell and display the results

![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/CustomCommand_Input.png)
![alt tag](https://github.com/bwya77/O365-Administration-Center/blob/master/Screenshots/CustomCommand_Output.png)

