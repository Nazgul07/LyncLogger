# Lync Logger

Logs text and audio conversations on Microsoft Lync & Skype for Business

## Description

Start the program to begin to log your Lync conversations.
This program adds an icon in the system tray bar.
Icon is green if logger is active. To activate the logger, just connect to Lync or Skype for Business.
Right-click on the notification tray icon brings up a menu. You can access the history folder from there or close the program.
![alt text](https://raw.githubusercontent.com/Nazgul07/LyncLogger/master/screenshot.png "ScreenShot")

## Office 365 Integration
You can enter your Office 365 Credentials to have your conversations saved to your Conversation History folder (useful if your organization disables Conversation History). Your password is stored Encrypted using the "System.Security.Cryptography.ProtectedData" class.


## Download

Download the latest version of LyncLogger here:
https://github.com/Nazgul07/LyncLogger/releases


## Notes

Tested on Skype for Business 2016 but should be compatible with Lync 2013.

To start the program automaticaly at startup you can put a shorcut in the folder "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"

Log files are located to "%appdata%\Lync logs" folder.


## How to contribute

This project depends on Microsoft dll:
- Microsoft.Lync.Model.dll
- Microsoft.Office.Uc.dll

Both are available at:
- http://www.microsoft.com/en-us/download/details.aspx?id=36824 (Lync 2013 & Skype for Business)

