<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
	 Created on:   	3/18/2019 11:17 AM
	 Created by:   	tausifkhan
	 Organization: 	FICO
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		Upgrade installed winscp version.
#>

Function Import-SMSTSENV
{
	try
	{
		$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
		Write-Output "$ScriptName - tsenv is $tsenv "
		$MDTIntegration = "YES"
		
		#$tsenv.GetVariables() | % { Write-Output "$ScriptName - $_ = $($tsenv.Value($_))" }
	}
	catch
	{
		Write-Output "$ScriptName - Unable to load Microsoft.SMS.TSEnvironment"
		Write-Output "$ScriptName - Running in standalonemode"
		$MDTIntegration = "NO"
	}
	Finally
	{
		if ($MDTIntegration -eq "YES")
		{
			$Logpath = $tsenv.Value("LogPath")
			$LogFile = $Logpath + "\" + "$ScriptName" + "$(get-date -format `"yyyyMMdd_hhmmsstt`").log"
			
		}
		Else
		{
			$Logpath = $env:TEMP
			$LogFile = $Logpath + "\" + "$ScriptName" + "$(get-date -format `"yyyyMMdd_hhmmsstt`").log"
		}
	}
}
Function Start-Logging
{
	start-transcript -path $LogFile -Force
}
Function Stop-Logging
{
	Stop-Transcript
}

# Set Vars
$SCRIPTDIR = split-path -parent $MyInvocation.MyCommand.Path
$SCRIPTNAME = split-path -leaf $MyInvocation.MyCommand.Path
$SOURCEROOT = "$SCRIPTDIR\Source"
$LANG = (Get-Culture).Name
$OSV = $Null
$ARCHITECTURE = $env:PROCESSOR_ARCHITECTURE

#Try to Import SMSTSEnv
. Import-SMSTSENV

#Start Transcript Logging
. Start-Logging

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Output base info
Write-Output ""
Write-Output "$ScriptName - ScriptDir: $ScriptDir"
Write-Output "$ScriptName - SourceRoot: $SOURCEROOT"
Write-Output "$ScriptName - ScriptName: $ScriptName"
Write-Output "$ScriptName - ScriptVersion: 988.1"
Write-Output "$ScriptName - Log: $LogFile"
###############


function get-latestwinscpversion
{
	
	$releases = 'https://winscp.net/eng/downloads.php'
	$re = 'WinSCP.+\.exe$'
	$download_page = Invoke-WebRequest -Uri $releases -UseBasicParsing
	
	$url = @($download_page.links | ? href -match $re) -notmatch 'beta|rc' | % href
	$url = 'https://winscp.net/eng' + $url
	$version = $url -split '-' | select -Last 1 -Skip 1
	$file_name = $url -split '/' | select -last 1
	return $version
}



function download_winscp
{
	Write-Output "***************Beginning function : downloadwinscp***********"
	$Path = "$env:Temp\"
	
	$releases = 'https://winscp.net/eng/downloads.php'
	$re = 'WinSCP.+\.exe$'
	$download_page = Invoke-WebRequest -Uri $releases -UseBasicParsing
	
	$DownloadUrl = @($download_page.links | ? href -match $re) -notmatch 'beta|rc' | % href
	$DownloadUrl = 'https://winscp.net' + $DownloadUrl
	$Results = Invoke-WebRequest -Method Get -Uri $DownloadUrl -MaximumRedirection 0 -UseBasicParsing -ErrorAction SilentlyContinue
	$winscpdllink = ($Results.Links | where { $_.class -eq 'btn btn-primary btn-lg' }).href
	$version = $DownloadUrl -split '-' | select -Last 1 -Skip 1
	$file_name = $DownloadUrl -split '/' | select -last 1
	$URL64 = "https://sourceforge.net/projects/winscp/files/WinSCP/$version/$file_name/download"
	@{
		Version		    = $version
		URL32		    = "https://sourceforge.net/projects/winscp/files/WinSCP/$version/$file_name/download"
		FileName32	    = $file_name
		ReleaseNotes    = "https://winscp.net/download/WinSCP-${version}-ReadMe.txt"
		FileType	    = 'exe'
		DownloadLInk    = $DownloadUrl
	}
	$Path = $Path + $file_name
	Write-Output "Installer save path is $Path"
	# Download the latest installer from box
	
	Write-Output "Downloading $file_name."
	#Invoke-WebRequest $DownloadUrl -OutFile $completepath
	$WebClient = New-Object System.Net.WebClient
	$WebClient.DownloadFile($winscpdllink, $Path)
	Rename-Item $Path "winscpsetup.exe" -Force -ErrorAction SilentlyContinue
	Start-Sleep -s 20
}

function Get-UninstallRegistryKey
{
	
	[CmdletBinding()]
	param (
		[parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
		[string]$softwareName,
		[parameter(ValueFromRemainingArguments = $true)]
		[Object[]]$ignoredArguments
	)
	
	#Write-FunctionCallLogMessage -Invocation $MyInvocation -Parameters $PSBoundParameters
	
	if ($softwareName -eq $null -or $softwareName -eq '')
	{
		throw "$SoftwareName cannot be empty for Get-UninstallRegistryKey"
	}
	
	$ErrorActionPreference = 'Stop'
	$local_key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
	$machine_key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
	$machine_key6432 = 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
	
	Write-Output "Retrieving all uninstall registry keys"
	[array]$keys = Get-ChildItem -Path @($machine_key6432, $machine_key, $local_key) -ErrorAction SilentlyContinue
	Write-Debug "Registry uninstall keys on system: $($keys.Count)"
	
	#Write-Output "Error handling check: `'Get-ItemProperty`' fails if a registry key is encoded incorrectly."
	[int]$maxAttempts = $keys.Count
	for ([int]$attempt = 1; $attempt -le $maxAttempts; $attempt++)
	{
		[bool]$success = $false
		
		$keyPaths = $keys | Select-Object -ExpandProperty PSPath
		try
		{
			[array]$foundKey = Get-ItemProperty -Path $keyPaths -ErrorAction Stop | ? { $_.DisplayName -match $softwareName }
			$success = $true
		}
		catch
		{
			Write-Debug "Found bad key."
			foreach ($key in $keys)
			{
				try
				{
					Get-ItemProperty $key.PsPath > $null
				}
				catch
				{
					$badKey = $key.PsPath
				}
			}
			Write-Output "Skipping bad key: $badKey"
			[array]$keys = $keys | ? { $badKey -NotContains $_.PsPath }
		}
		
		if ($success) { break; }
		
		if ($attempt -ge 10)
		{
			Write-Output "Found 10 or more bad registry keys. Run command again with `'--verbose --debug`' for more info."
			Write-Output "Each key searched should correspond to an installed program. It is very unlikely to have more than a few programs with incorrectly encoded keys, if any at all. This may be indicative of one or more corrupted registry branches."
		}
	}
	if ($foundKey -eq $null -or $foundkey.Count -eq 0)
	{
		Write-Output "No registry key found based on  '$softwareName'"
	}
	Write-Output "Found $($foundKey.Count) uninstall registry key(s) with SoftwareName:`'$SoftwareName`'";
	return $foundKey
}

Set-Alias Get-InstallRegistryKey Get-UninstallRegistryKey

Write-Output "Main starting"

Write-Output "Check if winscp is already installed"
$winscpinstalled = get-UninstallRegistryKey -SoftwareName "WinSCP" #| Select-Object -First 1
$installedversion = [System.Version]($winscpinstalled.DisplayVersion)

Write-Output "Installed version of winscp was determined to be $installedversion"

Write-Output "Trying to determine latest available version of WinSCP"
$version = [System.Version](get-latestwinscpversion)
Write-Output "Latest Version available of WinSCP is $version"

$Path = "$env:Temp\"
$Path = $Path + "winscpsetup.exe"

Write-Output "Path to winscp setup file is $Path"

if ($installedversion)
{
	if ($installedversion -ne $version)
	{
		Write-Output "WinSCP version $installedversion installed already, downloading newer version $version"
		download_winscp
		if (Test-Path $Path)
		{
			Write-Output "Winscp setup file present beginning install"
			$process = Start-Process $Path -ArgumentList "/VERYSILENT /NORESTART" -Wait -PassThru
			Write-Output "WinSCP setup exitcode: $($process.ExitCode)"
		}
		else
		{
			Write-Output "Download failed"
		}
		Remove-Item $Path -Force
	}
	elseif ($installedversion -eq $version)
	{
		Write-Output "Latest Version of WinSCP installed"
	}
}
elseif ($installedversion -eq $Null)
{
	Write-Output "WinSCP not installed"
}

#initiate hardware inventory 
Write-Output "Initiating Full Hardware inventory"
Get-WmiObject -Namespace root\ccm\invagt -Class InventoryActionStatus -Filter { InventoryActionID = '{00000000-0000-0000-0000-000000000001}' } | Remove-WmiObject
$SMSwmi = [wmiclass]"\root\ccm:SMS_Client"
$SMSwmi.TriggerSchedule("{00000000-0000-0000-0000-000000000001}")
Write-Output "Hardware inventory complete."


. Stop-Logging

