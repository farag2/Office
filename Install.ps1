# Install
if (Test-Path $PSScriptRoot\Office\Data\*\stream.x64.x-none.dat)
{
	Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/configure `"$PSScriptRoot\Config.xml`"" -Wait
}
else
{
	Write-Verbose -Message "There aren't neccessary Office files to install" -Verbose
}
