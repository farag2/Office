[![ko-fi](https://www.ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/Q5Q51QUJC)

## Overview

`Office.ps1` is a PowerShell script for Office 2016/2019 setup

Download and install `Office 2019` via ODT with pre-configured xml configurations

`Office 2019` supports `Windows 10` only

## Features

<details>
  <summary>List</summary>

 * General
  * Remove diagnostics tracking scheduled tasks
  * Do not send additional diagnostic and usage data to Microsoft
  * Disable LinkedIn features in Office applications
  * Turn off the cloud features
  * Turn on Touch/Mouse Mode

* Word
  * Do not show the Start screen when application starts
  * Do not open e-mail attachments and other uneditable files in reading view
  * Disable Protected View for files originating from the Internet
  * Disable Protected View for files located in potentially unsafe locations
  * Disable Protected View for Outlook attachments
  * Show the ruler
  * Save AutoRecover information every 3 minutes
  * Enable the "Draw" tab
  * Enable the "Developer" tab
  * Remove Adobe Acrobat Pro DC COM Add-ins

* Excel
  * Do not show the Start screen when application starts
  * Disable Protected View for files originating from the Internet
  * Disable Protected View for files located in potentially unsafe locations
  * Disable Protected View for Outlook attachments
  * Save AutoRecover information every 3 minutes
  * Enable the "Draw" tab
  * Enable the "Developer" tab
</details>
 
## Download and customize Office 2019

* E — Excel;
* O — Outlook;
* P — PowerPoint;
* W — Word.

* [Download](https://github.com/farag2/Office/releases) the archive from the release page and run `EOPW.cmd` or `EOPW.ps1` from the `Download` folder to download the whole Office 2019 package. It will be downloaded into the root folder (`Office`)
* After downloading run one of the install script **not as Administrator** from the `Install` folder

## Channels

* Monthly Channel, Standart
  * [Excel, Outlook, PowerPoint, and Word](https://github.com/farag2/Office/blob/master/XML/Download/EOPW.xml)
  * [Excel, Outlook, and Word](https://github.com/farag2/Office/blob/master/XML/Download/EOW.xml)
  * [Excel, PowerPoint, and Word](https://github.com/farag2/Office/blob/master/XML/Download/EPW.xml)
  * [Excel, Word](https://github.com/farag2/Office/blob/master/XML/Download/EW.xml)

## Links

* [Configure Office](https://config.office.com/deploymentsettings)
* [Overview of update channels](https://docs.microsoft.com/ru-ru/DeployOffice/overview-of-update-channels-for-office-365-proplus)
* [Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
* [Deploy Office](https://docs.microsoft.com/en-us/deployoffice/reference-articles-for-deploying-office-365-proplus)
* [Uninstall Office](https://support.microsoft.com/help/4027149)
  * SaRA
    * [zip](https://www.microsoft.com/en-us/download/100607)
    * [exe](https://aka.ms/SaRASetup)
  * [OffScrubC2R.vbs 2.15](https://github.com/farag2/Office/blob/master/Office%20Uninstall)
