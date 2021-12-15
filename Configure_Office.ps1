#region Features
# Remove diagnostics tracking scheduled tasks
Unregister-ScheduledTask -TaskName OfficeTelemetryAgentFallBack2016, OfficeTelemetryAgentLogOn2016 -Confirm:$false -ErrorAction Ignore

# Do not send additional diagnostic and usage data to Microsoft
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\Common\ClientTelemetry -Name SendTelemetry -PropertyType DWord -Value 3 -Force

# Disable LinkedIn features in Office applications
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn -Name OfficeLinkedIn -PropertyType DWord -Value 0 -Force

# Turn on Touch/Mouse Mode
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name OverridePointerMode -PropertyType DWord -Value 2 -Force
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name OverrideTabletMode -PropertyType DWord -Value 1 -Force

# Enable the dark theme
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name "UI Theme" -PropertyType DWord -Value 3 -Force
#endregion Features

#region Word
# Do not show the Start screen when application starts
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DisableBootToOfficeStart -PropertyType DWord -Value 1 -Force

# Do not open e-mail attachments and other uneditable files in reading view
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name AllowAutoReadingMode -PropertyType DWord -Value 0 -Force

# Disable Protected View for files originating from the Internet
if (-not (Test-Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Name DisableInternetFilesInPV -PropertyType DWord -Value 1 -Force

# Disable Protected View for files located in potentially unsafe locations
if (-not (Test-Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Name DisableUnsafeLocationsInPV -PropertyType DWord -Value 1 -Force

# Disable Protected View for Word attachments
if (-not (Test-Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Name DisableAttachmentsInPV -PropertyType DWord -Value 1 -Force

# Show the ruler
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name Ruler -PropertyType DWord -Value 1 -Force

# Save AutoRecover information every 3 minutes
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name AutosaveInterval -PropertyType DWord -Value 3 -Force

# Enable the "Draw" tab
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DrawInkTab -PropertyType DWord -Value 1 -Force

# Enable the "Developer" tab
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DeveloperTools -PropertyType DWord -Value 1 -Force
#endregion Word

#region Excel
# Do not show the Start screen when application starts
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name DisableBootToOfficeStart -PropertyType DWord -Value 1 -Force

# Save AutoRecover information every 3 minutes
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name AutoRecoverTime -PropertyType DWord -Value 3 -Force

# Enable the "Draw" tab
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name DrawInkTab -PropertyType DWord -Value 1 -Force

# Enable the "Developer" tab
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name DeveloperTools -PropertyType DWord -Value 1 -Force

# Maximaze the ribbon
if (-not (Test-Path -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Toolbars\Excel))
{
	New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Toolbars\Excel -Force
}
New-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Toolbars\Excel -Name QuickAccessToolbarStyle -PropertyType DWord -Value 16 -Force
#endregion Excel

#region Word
# Enable the "Developer" tab
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DeveloperTools -PropertyType DWord -Value 1 -Force

# Do not show the Start screen when application starts
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DisableBootToOfficeStart -PropertyType DWord -Value 1 -Force

# Show the ruler
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name Ruler -PropertyType DWord -Value 1 -Force

# Save AutoRecover information every 3 minutes
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name AutosaveInterval -PropertyType DWord -Value 3 -Force

# Enable the "Draw" tab
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DrawInkTab -PropertyType DWord -Value 1 -Force

# Maximaze the ribbon
if (-not (Test-Path -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Toolbars\Excel))
{
	New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Toolbars\Excel -Force
}
New-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Toolbars\Excel -Name QuickAccessToolbarStyle -PropertyType DWord -Value 16 -Force
#endregion Word

#region Outlook
# Do not show the Start screen when application starts
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Options -Name DisableBootToOfficeStart -PropertyType DWord -Value 1 -Force

# Enable the "Draw" tab
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Options -Name DrawInkTab -PropertyType DWord -Value 1 -Force

# Enable the classic ribbon
if (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences -Name EnableSingleLineRibbon -PropertyType DWord -Value 0 -Force
#endregion Outlook
