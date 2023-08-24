# Powershell script for Zabbix agents.

# Version 2.2 - for Zabbix agent 5x

## This script will check for pending Windows Updates, report them to Zabbix, and optionally install the updates (disabled by default).

### If you want the script to install updates, look for the comment in the script that tells you how to enable that function.

#### Check https://github.com/gitBalla/zabbix-winupdates for the latest version of this script from the repository forked from /spookoz/zabbix-winupdates

# ------------------------------------------------------------------------- #
# Variables
# ------------------------------------------------------------------------- #

# Change $reportpath to wherever you want your update reports to go.

$reportpath = "C:\zabbix\logs\Windows-Updates"

# Change $ZabbixInstallPath to wherever your Zabbix Agent is installed

$ZabbixInstallPath = "C:\zabbix\bin"
$ZabbixConfFile = "C:\zabbix\conf"

# Do not change the following variables unless you know what you are doing

$htReplace = New-Object hashtable
foreach ($letter in (Write-Output ä ae ö oe ü ue Ä Ae Ö Oe Ü Ue ß ss)) {
    $foreach.MoveNext() | Out-Null
    $htReplace.$letter = $foreach.Current
}
$pattern = "[$(-join $htReplace.Keys)]"
$returnStateOK = 0
$returnStateWarning = 1
$returnStateCritical = 2
$returnStateUnknown = 3
$returnStateOptionalUpdates = $returnStateWarning
$sender = "$ZabbixInstallPath\zabbix_sender.exe"
$senderArg1 = '-vv'
$senderArg2 = '-c'
$senderArg3 = "$ZabbixConfFile\zabbix_agent2.conf"
$senderArg4 = '-i'
$senderArgUpdateReboot = '\updateReboot.txt'
$senderArgLastUpdated = '\lastUpdated.txt'
$senderArgCountCritical = '\countCritical.txt'
$senderArgCountImportant = '\countImportant.txt'
$senderArgCountOptional = '\countOptional.txt'
$senderArgCountHidden = '\countHidden.txt'
$CountCriticalNum = '\CountCriticalNum.txt'
$senderArg5 = '-k'
$senderArgUpdating = 'Windows-Updates.Updating'
$senderArg6 = '-o'
$senderArg7 = '0'
$senderArg8 = '1'
$token="Hostname"
$extractedValue = Get-Content $senderArg3
$hostname = (($extractedValue -split [System.Environment]::NewLine) | where {$_ -Like "$token*"}).Substring("$token=".Length);


If(!(test-path $reportpath))
{
      New-Item -ItemType Directory -Force -Path $reportpath
}

# ------------------------------------------------------------------------- #
# This part gets the date Windows Updates were last applied and writes it to temp file
# ------------------------------------------------------------------------- #

$windowsUpdateObject = New-Object -ComObject Microsoft.Update.AutoUpdate
Write-Output "- Windows-Updates.LastUpdated $($windowsUpdateObject.Results.LastInstallationSuccessDate)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgLastUpdated

# ------------------------------------------------------------------------- #
# This part get the reboot status and writes to test file
# ------------------------------------------------------------------------- #

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"){ 
	Write-Output "- Windows-Updates.Reboot 1" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgUpdateReboot
    Write-Host "`t There is a reboot pending" -ForeGroundColor "Red"
}else {
	Write-Output "- Windows-Updates.Reboot 0" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgUpdateReboot
    Write-Host "`t No reboot pending" -ForeGroundColor "Green"
		}
# ------------------------------------------------------------------------- #		
# This part checks available Windows updates
# ------------------------------------------------------------------------- #

$updateSession = new-object -com "Microsoft.Update.Session"
$updates=$updateSession.CreateupdateSearcher().Search(("IsInstalled=0 and Type='Software'")).Updates

$criticalTitles = "";
$importantTitles = "";
$countCritical = 0;
$countImportant = 0;
$countOptional = 0;
$countHidden = 0;

# ------------------------------------------------------------------------- #
# If no updates are required - it writes the info to a temp file, sends it to Zabbix server and exits
# ------------------------------------------------------------------------- #

if ($updates.Count -eq 0) {

	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$CountCriticalNum
	Write-Output "- Windows-Updates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountCritical
	Write-Output "- Windows-Updates.Important $($countImportant)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountImportant
	Write-Output "- Windows-Updates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountOptional
	Write-Output "- Windows-Updates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountHidden
    Write-Host "`t There are no pending updates" -ForeGroundColor "Green"
	
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgUpdateReboot -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgLastUpdated -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountCritical -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountImportant -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountOptional -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountHidden -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg5 $senderArgUpdating $senderArg6 $senderArg7 -s "$hostname"
	
	exit $returnStateOK
}

# ------------------------------------------------------------------------- #
# This part counts the number of updates to be applied
# ------------------------------------------------------------------------- #

foreach ($update in $updates) {
	if ($update.IsHidden) {
		$countHidden++
	} elseif ($update.AutoSelectOnWebSites -and $update.MsrcSeverity -eq "Critical") {
		$criticalTitles += $update.Title + " `n"
		$countCritical++
	} elseif ($update.AutoSelectOnWebSites -and $update.MsrcSeverity -eq "Important") {
		$importantTitles += $update.Title + " `n"
		$countImportant++
	} else {
		$countOptional++
	}
}

# ------------------------------------------------------------------------- #
# This part writes the number of each update required to a temp file and sends it to Zabbix
# ------------------------------------------------------------------------- #

if (($countCritical + $countImportant + $countOptional) -gt 0) {

	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$CountCriticalNum
	Write-Output "- Windows-Updates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountCritical
	Write-Output "- Windows-Updates.Important $($countImportant)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountImportant
	Write-Output "- Windows-Updates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountOptional
	Write-Output "- Windows-Updates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountHidden
    Write-Host "`t There are $($countCritical) critical updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countImportant) important updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countOptional) optional updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countHidden) hidden updates available" -ForeGroundColor "Yellow"
	
    & $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgUpdateReboot -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgLastUpdated -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountCritical -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountImportant -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountOptional -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountHidden -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg5 $senderArgUpdating $senderArg6 $senderArg7 -s "$hostname"
}   

# ------------------------------------------------------------------------- #
# The following section will automatically apply any pending updates if it finds any critical updates missing or more than 3 optional updates missing. If you want this to run, remove the <# and #> comment block lines.

<#
if ($countCritical -gt 0 -Or $countImportant -gt 0 -Or $countOptional -gt 2) {
		
			& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg5 $senderArgupdating $senderArg6 $senderArg8
			$ErrorActionPreference = "SilentlyContinue"
			
			If ($Error) {
				$Error.Clear()
			}
			
			$Today = Get-Date
			$TodayFile = get-date -f yyyy-MM-dd
			$UpdateCollection = New-Object -ComObject Microsoft.Update.UpdateColl
			$Searcher = New-Object -ComObject Microsoft.Update.Searcher
			$Session = New-Object -ComObject Microsoft.Update.Session

			Write-Host "`t Initialising and Checking for Applicable Updates. Please wait ..." -ForeGroundColor "Yellow"
			$Result = $Searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

				$ReportFile = $reportpath + "\" + $hostname + "_Windows-UpdatesReport_"  + $TodayFile + ".txt"
				If (Test-Path $ReportFile) {
					Remove-Item $ReportFile
				}
				New-Item $ReportFile -Type File -Force -Value "Windows Update Report For Computer: $hostname`r`n" | Out-Null
				Add-Content $ReportFile "Report Created On: $Today`r"
				
			If ($Result.Updates.Count -EQ 0) {
				Write-Host "`t There are no applicable updates for this computer."
				Add-Content $ReportFile "==============================================================================`r`n"
				Add-Content $ReportFile "There are no applicable updates for this computer today.`r`n"
				Add-Content $ReportFile "------------------------------------------------`r"
			}
			Else {

				Add-Content $ReportFile "==============================================================================`r`n"
				Write-Host "`t Preparing List of Applicable Updates For This Computer ..." -ForeGroundColor "Yellow"
				Add-Content $ReportFile "List of Applicable Updates For This Computer`r"
				Add-Content $ReportFile "------------------------------------------------`r"
				For ($Counter = 0; $Counter -LT $Result.Updates.Count; $Counter++) {
					$DisplayCount = $Counter + 1
						$Update = $Result.Updates.Item($Counter)
					$UpdateTitle = $Update.Title
					Add-Content $ReportFile "`t $DisplayCount -- $UpdateTitle"
				}
				$Counter = 0
				$DisplayCount = 0
				Add-Content $ReportFile "`r`n"
				Write-Host "`t Initialising Download of Applicable Updates ..." -ForegroundColor "Yellow"
				Add-Content $ReportFile "Initialising Download of Applicable Updates"
				Add-Content $ReportFile "------------------------------------------------`r"
				$Downloader = $Session.CreateUpdateDownloader()
				$UpdatesList = $Result.Updates
				For ($Counter = 0; $Counter -LT $Result.Updates.Count; $Counter++) {
					$UpdateCollection.Add($UpdatesList.Item($Counter)) | Out-Null
					$ShowThis = $UpdatesList.Item($Counter).Title
					$DisplayCount = $Counter + 1
					Add-Content $ReportFile "`t $DisplayCount -- Downloading Update $ShowThis `r"
					$Downloader.Updates = $UpdateCollection
					$Track = $Downloader.Download()
					If (($Track.HResult -EQ 0) -AND ($Track.ResultCode -EQ 2)) {
						Add-Content $ReportFile "`t Download Status: SUCCESS"
					}
					Else {
						Add-Content $ReportFile "`t Download Status: FAILED With Error -- $Error()"
						$Error.Clear()
						Add-content $ReportFile "`r"
					}	
				}
				$Counter = 0
				$DisplayCount = 0
				Write-Host "`t Starting Installation of Downloaded Updates ..." -ForegroundColor "Yellow"
				Add-Content $ReportFile "`r`n"
				Add-Content $ReportFile "Installation of Downloaded Updates"
				Add-Content $ReportFile "------------------------------------------------`r"
				$Installer = New-Object -ComObject Microsoft.Update.Installer
				For ($Counter = 0; $Counter -LT $UpdateCollection.Count; $Counter++) {
					$Track = $Null
					$DisplayCount = $Counter + 1
					$WriteThis = $UpdateCollection.Item($Counter).Title
					Add-Content $ReportFile "`t $DisplayCount -- Installing Update: $WriteThis"
					$Installer.Updates = $UpdateCollection
					Try {
						$Track = $Installer.Install()
						Add-Content $ReportFile "`t Update Installation Status: SUCCESS"
					}
					Catch {
						[System.Exception]
						Add-Content $ReportFile "`t Update Installation Status: FAILED With Error -- $Error()"
						$Error.Clear()
						Add-content $ReportFile "`r"
					}	
				}
			}


	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg5 $senderArgupdating $senderArg6 $senderArg7


    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"){ 
	    Write-Output "- Windows-Updates.Reboot 1" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgUpdateReboot
        Write-Host "`t There is a reboot pending" -ForeGroundColor "Red"
    }else {
	    Write-Output "- Windows-Updates.Reboot 0" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgUpdateReboot
        Write-Host "`t No reboot pending" -ForeGroundColor "Green"
		    }

    $updates=$updateSession.CreateupdateSearcher().Search(("IsInstalled=0 and Type='Software'")).Updates

    Write-Output "- Windows-Updates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountCritical
    Write-Output "- Windows-Updates.Important $($countImportant)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountImportant
	Write-Output "- Windows-Updates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountOptional
	Write-Output "- Windows-Updates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountHidden
    Write-Host "`t There are now $($countCritical) critical updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are now $($countOptional) optional updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are now $($countHidden) hidden updates available" -ForeGroundColor "Yellow"

    & $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgUpdateReboot -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgLastUpdated -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountCritical -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountOptional -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountHidden -s "$hostname"

	exit $returnStateCritical
}
#>

# Remove the comment block lines above here to enable automatic updating
# ------------------------------------------------------------------------- #

if ($countOptional -gt 0) {
	exit $returnStateOptionalUpdates
}

# ------------------------------------------------------------------------- #
# If Hidden Updates are found, this part will write the info to a temp file, send to Zabbix server and exit
# ------------------------------------------------------------------------- #

if ($countHidden -gt 0) {
	
	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$CountCriticalNum
	Write-Output "- Windows-Updates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountCritical
	Write-Output "- Windows-Updates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountOptional
	Write-Output "- Windows-Updates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$senderArgCountHidden
    Write-Host "`t There are $($countCritical) critical updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countOptional) optional updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countHidden) hidden updates available" -ForeGroundColor "Yellow"
	
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgUpdateReboot -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgLastUpdated -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountCritical -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountOptional -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg4 $env:temp$senderArgCountHidden -s "$hostname"
	& $sender $senderArg1 $senderArg2 $senderArg3 $senderArg5 $senderArgUpdating $senderArg6 $senderArg7 -s "$hostname"
	
	exit $returnStateOK
}

exit $returnStateUnknown