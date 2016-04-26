# O365 Administration Center
![alt tag](http://www.gnu.org/graphics/gplv3-88x31.png)

v2.0.2

[.exe download] (https://github.com/bwya77/O365-Administration-Center/blob/master/O365%20Administration%20Center.exe?raw=true)

[.msi download] (https://github.com/bwya77/O365-Administration-Center/blob/master/O365%20Administration%20Center.msi?raw=true)

___

##Description
The O365 Administration Center is an application written in PowerShell that lets administrators easily and quickly manage their Office 365 environment. It allows partner accounts to connect to all of their tenants and run the same commands. The results are shown in the textbox that also accepts custom commands as input. The commands can be typed in, and by pressing the enter key or the “Run Command” button the command is passed through to PowerShell and the results are displayed back on the same textbox. Results can also be exported to a file easily by using the “Export to File” button which uses the Out-File cmdlet.


Included is the .exe if you just want to run it, or the .msi if you want to install it.
___

##Prerequisites

Windows Versions:

Windows 7 Service Pack 1 (SP1) or higher

Windows Server 2008 R2 SP1 or higher

You need to install the Microsoft.NET Framework 4.5 or later and then Windows Management Framework 3.0 or later. 

MS-Online module needs to be installed. Install the MSOnline Services Sign In Assistant: https://www.microsoft.com/en-us/download/details.aspx?id=41950 

Azure Active Directory Module for Windows PowerShell needs to be installed: http://go.microsoft.com/fwlink/p/?linkid=236297

The Microsoft Online Services Sign-In Assistant provides end user sign-in capabilities to Microsoft Online Services, such as Office 365.

Windows PowerShell needs to be configured to run scripts, and by default, it isn't. To enable Windows PowerShell to run scripts, run the following command in an elevated Windows PowerShell window (a Windows PowerShell window you open by selecting Run as administrator):
Set-ExecutionPolicy Unrestricted"

PowerShell v3 or higher
___

##.Screenshots

####O365 GUI
![alt tag](http://i.imgur.com/X5ERaSG.png?1)

####Login to O365
You are prompted for O365 credentials. It will then load all Exchange Online cmdlets.
![alt tag](http://i.imgur.com/yRj2pj5.png)

####Partner List
![alt tag](http://i.imgur.com/svxIibW.png)

