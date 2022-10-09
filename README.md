[![ko-fi](https://www.ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/Q5Q51QUJC)

## How-to

* Choose which Offce to download by editing the `DownloadOffice` arguments in the of the file

  ```powershell
    DownloadOffice -Branch ProPlus2019Retail -Channel Current -Components Word, Excel, PowerPoint
    DownloadOffice -Branch ProPlus2021Volume -Channel PerpetualVL2021 -Components Excel, Word
    DownloadOffice -Branch O365ProPlusRetail -Channel BetaChannel -Components Excel, OneDrive, Outlook, PowerPoint, Teams, Word
  ```

* You may uncomment this string in the file to install Office automatically after it's downloaded (the script downloads Office by default only)

  ```powershell
  Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/configure `"$PSScriptRoot\Config.xml`"" -Wait
  ```

## Addendum

`Configure_Office.ps1` is a script for configuring Office 2016/2019/365
`Office 2019, 2021, & 365` support `Windows 10` & `Windows 11` only

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

## Links

* [Configure Office](https://config.office.com/deploymentsettings)
* [Overview of update channels](https://docs.microsoft.com/en-us/DeployOffice/overview-of-update-channels-for-office-365-proplus)
* [Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
* [Deploy Office](https://docs.microsoft.com/en-us/deployoffice/reference-articles-for-deploying-office-365-proplus)
* [Uninstall Office (SaRA)](https://www.microsoft.com/en-us/download/100607)
* [OffScrubC2R.vbs 2.19](https://github.com/farag2/Office/tree/master/Office_Uninstall)
* [Office Tool Plus](https://github.com/YerongAI/Office-Tool)
