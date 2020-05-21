## Overview
PowerShell script for Office 2016/2019 setup

## Script Features
- General
    - Remove diagnostics tracking scheduled tasks
    - Do not send additional diagnostic and usage data to Microsoft
    - Disable LinkedIn features in Office applications
    - Turn off the cloud features
    - Turn on Touch/Mouse Mode
- Word
  - Do not show the Start screen when application starts
  - Do not open e-mail attachments and other uneditable files in reading view
  - Disable Protected View for files originating from the Internet
  - Disable Protected View for files located in potentially unsafe locations
  - Disable Protected View for Outlook attachments
  - Show the ruler
  - Save AutoRecover information every 3 minutes
  - Enable the "Draw" tab
  - Enable the "Developer" tab
  - Remove Adobe Acrobat Pro DC COM Add-ins
- Excel
  - Do not show the Start screen when application starts
  - Disable Protected View for files originating from the Internet
  - Disable Protected View for files located in potentially unsafe locations
  - Disable Protected View for Outlook attachments
  - Save AutoRecover information every 3 minutes
  - Enable the "Draw" tab
  - Enable the "Developer" tab

## Customized .xml configs

XML Configurations for downloading and installing Office 2019
 - E — Excel;
 - O — Outlook;
 - P — PowerPoint;
 - W — Word.

Place .xml in XML folder, and setup.exe with .cmd (download/install) in a directory upper than "XML".
Run .cmd (download/install) **not as Administrator**

## Channels
- VL, Semi-Annual Channel, PerpetualVL2019
   - [Excel, Outlook, PowerPoint, and Word](https://github.com/farag2/Office/blob/master/XML/EOPW_VL.xml)
   - [Excel, Outlook,and Word](https://github.com/farag2/Office/blob/master/XML/EOW_VL.xml)
   - [Excel, PowerPoint, and Word](https://github.com/farag2/Office/blob/master/XML/EPW_VL.xml)
   - [Excel, Word](https://github.com/farag2/Office/blob/master/XML/EW_VL.xml)

- Monthly Channel, Standart
   - [Excel, Outlook, PowerPoint, and Word](https://github.com/farag2/Office/blob/master/XML/EOPW.xml)
   - [Excel, Outlook, and Word](https://github.com/farag2/Office/blob/master/XML/EOW.xml)
   - [Excel, PowerPoint, and Word](https://github.com/farag2/Office/blob/master/XML/EPW.xml)
   - [Excel, Word](https://github.com/farag2/Office/blob/master/XML/EW.xml)

## Links
- [Configure Office](https://config.office.com/deploymentsettings)
- [Overview of update channels](https://docs.microsoft.com/ru-ru/DeployOffice/overview-of-update-channels-for-office-365-proplus)
- [Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
- [Deploy Office](https://docs.microsoft.com/en-us/deployoffice/reference-articles-for-deploying-office-365-proplus)
- [Uninstall Office](https://support.microsoft.com/help/4027149)
   - [Latest version (2.15) of OffScrubC2R.vbs](https://github.com/farag2/Office/blob/master/Office%20Uninstall)
