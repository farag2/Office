<#
	.SYNOPSIS
	Download Office 2019 or 365

	.PARAMETER Branch
	Choose Office branch: 2019 or 365

	.PARAMETER Channel
	Choose Office channel: Current or SemiAnnual

	.PARAMETER Components
	Choose Office components: Access, OneDrive, Outlook, Word, Excel, PowerPoint, Teams

	.EXAMPLE Download Office 2019 with the Word, Excel, PowerPoint components
	Office -Branch 2019 -Channel Current -Components Word, Excel, PowerPoint

	.EXAMPLE Download Office 365 with the Word, Excel, PowerPoint components
	Office -Branch 365 -Channel SemiAnnual -Components Word, Excel, PowerPoint

	.LINK
	https://config.office.com/deploymentsettings
#>
function Office
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateSet("2019", "365")]
		[string]
		$Branch,

		[Parameter(Mandatory = $true)]
		[ValidateSet("Current", "SemiAnnual")]
		[string]
		$Channel,

		[Parameter(Mandatory = $true)]
		[ValidateSet("Access", "OneDrive", "Outlook", "Word", "Excel", "PowerPoint", "Teams")]
		[string[]]
		$Components
	)

	if (-not (Test-Path -Path "$PSScriptRoot\Default.xml"))
	{
		Write-Warning -Message "Default.xml doesn't exist"
		exit
	}

	[xml]$Config = Get-Content -Path "$PSScriptRoot\Default.xml" -Encoding Default -Force

	switch ($Branch)
	{
		2019
		{
			($Config.Configuration.Add.Product | Where-Object -FilterScript {$_.ID -eq ""}).ID = "Standard2019Retail"
		}
		365
		{
			($Config.Configuration.Add.Product | Where-Object -FilterScript {$_.ID -eq ""}).ID = "O365ProPlusRetail"
		}
	}

	switch ($Channel)
	{
		Current
		{
			($Config.Configuration.Add | Where-Object -FilterScript {$_.Channel -eq ""}).Channel = "Current"
		}
		SemiAnnual
		{
			($Config.Configuration.Add | Where-Object -FilterScript {$_.Channel -eq ""}).Channel = "SemiAnnual"
		}
	}

	foreach ($Component in $Components)
	{
		switch ($Component)
		{
			Access
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Access']")
				$Node.ParentNode.RemoveChild($Node)
			}
			Excel
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Excel']")
				$Node.ParentNode.RemoveChild($Node)
			}
			OneDrive
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='OneDrive']")
				$Node.ParentNode.RemoveChild($Node)
			}
			Outlook
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Outlook']")
				$Node.ParentNode.RemoveChild($Node)
			}
			Word
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Word']")
				$Node.ParentNode.RemoveChild($Node)
			}
			PowerPoint
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='PowerPoint']")
				$Node.ParentNode.RemoveChild($Node)
			}
			Teams
			{
				$Node = $Config.SelectSingleNode("//ExcludeApp[@ID='Teams']")
				$Node.ParentNode.RemoveChild($Node)
			}
		}
	}

	$Config.Save("$PSScriptRoot\Config.xml")

	# Download Office Deployement Tool
	# https://www.microsoft.com/en-us/download/details.aspx?id=49117
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

	$Parameters = @{
		Uri             = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
		UseBasicParsing = $true
		Verbose         = $true
	}
	$ODTURL = ((Invoke-WebRequest @Parameters).Links | Where-Object {$_.outerHTML -like "*click here to download manually*"}).href
	$Parameters = @{
		Uri             = $ODTURL
		OutFile         = "$PSScriptRoot\officedeploymenttool.exe"
		UseBasicParsing = $true
		Verbose         = $true
	}
	Invoke-WebRequest @Parameters

	# Expand officedeploymenttool.exe
	Start-Process "$PSScriptRoot\officedeploymenttool.exe" -ArgumentList "/quiet /extract:`"$PSScriptRoot\officedeploymenttool`"" -Wait

	$Parameters = @{
		Path        = "$PSScriptRoot\officedeploymenttool\setup.exe"
		Destination = "$PSScriptRoot"
		Force       = $true
	}
	Move-Item @Parameters

	Start-Sleep -Seconds 1

	Remove-item -Path "$PSScriptRoot\officedeploymenttool", "$PSScriptRoot\officedeploymenttool.exe"  -Recurse -Force

	# Start downloading to the Office folder
	Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/download","`"$PSScriptRoot\Config.xml`"" -Wait
}

Office -Branch 365 -Channel SemiAnnual -Components Word, Excel, PowerPoint

# Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/configure `"$PSScriptRoot\Config.xml`"" -Wait
