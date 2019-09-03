# Remove diagnostics tracking scheduled tasks
# Удалить задачи диагностического отслеживания
Unregister-ScheduledTask OfficeTelemetryAgentFallBack2016, OfficeTelemetryAgentLogOn2016 -Confirm:$false -ErrorAction SilentlyContinue
# Do not send additional diagnostic and usage data to Microsoft
# Выключить необязательные сетевые функции
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\Common\ClientTelemetry -Name SendTelemetry -Value 1 -Force
# Disable LinkedIn features in Office applications
# Не использовать функции LinkedIn
IF (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn -Name OfficeLinkedIn -Value 0 -Force
# Turn off the cloud features
# Отключить облачную интеграцию
IF (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Common\SignIn))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\SignIn -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\SignIn -Name SignInOptions -Value 3 -Force
# Turn on Touch/Mouse Mode
# Включить сенсорное управление
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name OverridePointerMode -Value 2 -Force
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name OverrideTabletMode -Value 1 -Force
# Word
# Do not show the Start screen when application starts
# Не показывать начальный экран при запуске
IF (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DisableBootToOfficeStart -Value 1 -Force
# Do not open e-mail attachments and other uneditable files in reading view
# Не открывать вложения электронной почты и другие нередактируемые файлы в режиме чтения и защищенный просмотр
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name AllowAutoReadingMode -Value 0 -Force
# Disable Protected View for files originating from the Internet
# Отключить защищенный просмотр для файлов из Интернета
IF (-not (Test-Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Name DisableInternetFilesInPV -Value 1 -Force
# Disable Protected View for files located in potentially unsafe locations
# Отключить защищенный просмотр для файлов в потенциально небезопасных расположениях
IF (-not (Test-Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Name DisableUnsafeLocationsInPV -Value 1 -Force
# Disable Protected View for Outlook attachments
# Отключить защищенный просмотр для вложений Outlook
IF (-not (Test-Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Security\ProtectedView -Name DisableAttachmentsInPV -Value 1 -Force
# Show the ruler
# Отобразить линейку
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name Ruler -Value 1 -Force
# Save AutoRecover information every 3 minutes
# Включить автосохранение каждые 3 минуты
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name AutosaveInterval -Value 3 -Force
# Enable the "Draw" tab
# Включить вкладку "Рисование"
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DrawInkTab -Value 1 -Force
# Enable the "Developer" tab
# Включить вкладку "Разработчик"
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DeveloperTools -Value 1 -Force
# Remove Adobe Acrobat Pro DC COM Add-ins
# Удалить надстройки COM Adobe Acrobat Pro DC
Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Office\Word\Addins\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Word\Addins\*" -Recurse -Force -ErrorAction SilentlyContinue
# Excel
# Do not show the Start screen when application starts
# Не показывать начальный экран при запуске
IF (-not (Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name DisableBootToOfficeStart -Value 1 -Force
# Do not open e-mail attachments and other uneditable files in reading view
# Отключить открытие вложения электронной почты и другие нередактируемые файлы в режиме чтения и защищенный просмотр
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name AllowAutoReadingMode -Value 0 -Force
# Disable Protected View for files originating from the Internet
# Отключить защищенный просмотр для файлов из Интернета
IF (-not (Test-Path HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView -Name DisableInternetFilesInPV -Value 1 -Force
# Disable Protected View for files located in potentially unsafe locations
# Отключить защищенный просмотр для файлов в потенциально небезопасных расположениях
IF (-not (Test-Path HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView -Name DisableUnsafeLocationsInPV -Value 1 -Force
# Disable Protected View for Outlook attachments
# Отключить защищенный просмотр для вложений Outlook
IF (-not (Test-Path HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView))
{
	New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView -Force
}
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView -Name DisableAttachmentsInPV -Value 1 -Force
# Save AutoRecover information every 3 minutes
# Включить автосохранение каждые 3 минуты
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name AutoRecoverTime -Value 3 -Force
# Enable the "Draw" tab
# Включить вкладку "Рисование"
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name DrawInkTab -Value 1 -Force
# Enable the "Developer" tab
# Включить вкладку "Разработчик"
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name DeveloperTools -Value 1 -Force