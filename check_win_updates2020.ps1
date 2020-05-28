# Powershell script for Zabbix agents.

# Version 1.1

## This script will check for pending Windows Updates, report them to Zabbix, and optionally install the updates.

### If you do not wish the script to install updates, look for the comment in the script that tells you how to disable that function.

#### Check https://github.com/SpookOz/zabbix-winupdates for the latest version of this script

# ------------------------------------------------------------------------- #
# Variables
# ------------------------------------------------------------------------- #

# Change $reportpath to wherever you want your update reports to go.

$reportpath = "C:\IT\WinUpdates"

# Change $ZabbixInstallPath to wherever your Zabbix Agent is installed

$ZabbixInstallPath = "$Env:Programfiles\Zabbix Agent"

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
$Sender = "$ZabbixInstallPath\zabbix_sender.exe"
$Senderarg1 = '-vv'
$Senderarg2 = '-c'
$Senderarg3 = "$ZabbixInstallPath\zabbix_agentd.conf"
$Senderarg4 = '-i'
$SenderargUpdateReboot = '\updatereboot.txt'
$Senderarglastupdated = '\lastupdated.txt'
$Senderargcountcritical = '\countcritical.txt'
$SenderargcountOptional = '\countOptional.txt'
$SenderargcountHidden = '\countHidden.txt'
$Countcriticalnum = '\countcriticalnum.txt'
$Senderarg5 = '-k'
$Senderargupdating = 'Winupdates.Updating'
$Senderarg6 = '-o'
$Senderarg7 = '0'
$Senderarg8 = '1'


If(!(test-path $reportpath))
{
      New-Item -ItemType Directory -Force -Path $reportpath
}

# ------------------------------------------------------------------------- #
# This part gets the date Windows Updates were last applied and writes it to temp file
# ------------------------------------------------------------------------- #

$windowsUpdateObject = New-Object -ComObject Microsoft.Update.AutoUpdate
Write-Output "- Winupdates.LastUpdated $($windowsUpdateObject.Results.LastInstallationSuccessDate)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderarglastupdated

# ------------------------------------------------------------------------- #
# This part get the reboot status and writes to test file
# ------------------------------------------------------------------------- #

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"){ 
	Write-Output "- Winupdates.Reboot 1" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargUpdateReboot
}else {
	Write-Output "- Winupdates.Reboot 0" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargUpdateReboot
		}
# ------------------------------------------------------------------------- #		
# This part checks available Windows updates
# ------------------------------------------------------------------------- #

$updateSession = new-object -com "Microsoft.Update.Session"
$updates=$updateSession.CreateupdateSearcher().Search(("IsInstalled=0 and Type='Software'")).Updates

$criticalTitles = "";
$countCritical = 0;
$countOptional = 0;
$countHidden = 0;

# ------------------------------------------------------------------------- #
# If no updates are required - it writes the info to a temp file, sends it to Zabbix server and exits
# ------------------------------------------------------------------------- #

if ($updates.Count -eq 0) {

	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$Countcriticalnum
	Write-Output "- Winupdates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderargcountcritical
	Write-Output "- Winupdates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountOptional
	Write-Output "- Winupdates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountHidden
	
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargUpdateReboot
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderarglastupdated
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderargcountcritical
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountOptional
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountHidden
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg7
	
	exit $returnStateOK
}

# ------------------------------------------------------------------------- #
# This part counts the number of updates to be applied
# ------------------------------------------------------------------------- #

foreach ($update in $updates) {
	if ($update.IsHidden) {
		$countHidden++
	}
	elseif ($update.AutoSelectOnWebSites) {
		$criticalTitles += $update.Title + " `n"
		$countCritical++
	} else {
		$countOptional++
	}
}

# ------------------------------------------------------------------------- #
# This part writes the number of each update required to a temp file and sends it to Zabbix
# ------------------------------------------------------------------------- #

if (($countCritical + $countOptional) -gt 0) {

	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$Countcriticalnum
	Write-Output "- Winupdates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderargcountcritical
	Write-Output "- Winupdates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountOptional
	Write-Output "- Winupdates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountHidden
	
    & $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargUpdateReboot
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderarglastupdated
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderargcountcritical
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountOptional
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountHidden
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg7
}   

# ------------------------------------------------------------------------- #
# The following section will automatically apply any pending updates if it finds any critical updates missing or more than 3 optional updates missing. If you do not want this to run, comment out or delete everything between here and the next comment.

	
if ($countCritical -gt 0 -Or $countOptional -gt 2) {
		
			& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg8
			$ErrorActionPreference = "SilentlyContinue"
			
			If ($Error) {
				$Error.Clear()
			}
			
			$Today = Get-Date
			$TodayFile = get-date -f yyyy-MM-dd
			$UpdateCollection = New-Object -ComObject Microsoft.Update.UpdateColl
			$Searcher = New-Object -ComObject Microsoft.Update.Searcher
			$Session = New-Object -ComObject Microsoft.Update.Session

			Write-Host
			Write-Host "`t Initialising and Checking for Applicable Updates. Please wait ..." -ForeGroundColor "Yellow"
			$Result = $Searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

				$ReportFile = $reportpath + "\" + $Env:ComputerName + "_WinupdateReport_"  + $TodayFile + ".txt"
				If (Test-Path $ReportFile) {
					Remove-Item $ReportFile
				}
				New-Item $ReportFile -Type File -Force -Value "Windows Update Report For Computer: $Env:ComputerName`r`n" | Out-Null
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


	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg7


    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"){ 
	    Write-Output "- Winupdates.Reboot 1" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargUpdateReboot
    }else {
	    Write-Output "- Winupdates.Reboot 0" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargUpdateReboot
		    }

    $updates=$updateSession.CreateupdateSearcher().Search(("IsInstalled=0 and Type='Software'")).Updates

    Write-Output "- Winupdates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderargcountcritical
	Write-Output "- Winupdates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountOptional
	Write-Output "- Winupdates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountHidden

    & $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargUpdateReboot
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderarglastupdated
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderargcountcritical
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountOptional
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountHidden

	exit $returnStateCritical
}

# Comment out or delete everything above here to disable automatic updating
# ------------------------------------------------------------------------- #

if ($countOptional -gt 0) {
	exit $returnStateOptionalUpdates
}

# ------------------------------------------------------------------------- #
# If Hidden Updates are found, this part will write the info to a temp file, send to Zabbix server and exit
# ------------------------------------------------------------------------- #

if ($countHidden -gt 0) {
	
	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$Countcriticalnum
	Write-Output "- Winupdates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderargcountcritical
	Write-Output "- Winupdates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountOptional
	Write-Output "- Winupdates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountHidden
	
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargUpdateReboot
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderarglastupdated
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderargcountcritical
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountOptional
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountHidden
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg7
	
	exit $returnStateOK
}

exit $returnStateUnknown