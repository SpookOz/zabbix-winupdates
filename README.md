# zabbix-winupdates

This is a Zabbix template to monitor and (optionally) run Windows updates for active Windows agents. It has been tested on Zabbix 4.4.8 to v6.4 with Windows 8.1, 10, server 2019, and 2022. Currently needs testing with other Windows versions and other Zabbix versions.

## Features

- Works with active agents (behind NAT - so perfect remote sites).
- Checks for critical updates, important updates, optional updates and hidden updates and reports numbers of each.
- Reports date of last updates.
- Reports if a reboot is pending.
- If there is a critical or important update pending or 3 or more optional updates, runs Windows Update to patch machine and reports back that updates are running (this can be disabled).
- As part of above it writes a report file to C:\IT\Winupdates by default. This can be changed in the Powershell Script
- Includes a panel for the Dashboard.
- Includes triggers to warn for different update states.
- Includes an option to auto-download the Powershell script so you don't have to manually deploy it to your hosts (disabled by default).


## Requirements

- Has only been tested with Zabbix active agents. It will need tweaking to work with passive agents.
- At this stage it has only been tested on Windows 8.1, 10, Server 2019 and 2022. However, it should also work on server 2016 and Windows 7. Not sure about previous versions at this stage.
- If you are running <v5 Zabbix agent, you MUST have EnableRemoteCommands=1 in your conf file (so it can run the PS script).
- If you are running >v5 Zabbix agent, you MUST have AllowKey=system.run[*] in your conf file (https://www.zabbix.com/documentation/current/manual/config/notifications/action/operation/remote_command)
- This script assumes the Zabbix client is installed in the Program Files directory. You will need to modify the Powershell script if it is not.

## Files Included

Zabbix Agents >v5
- Winupdates.xml: Template to import into Zabbix.
- check_win_updates2020.ps1: Powershell script. Must be placed in a "plugins" directory under your Zabbix Agent install folder (eg: C:\Program Files\Zabbix Agent\plugins).

Zabbix Agents >v5
- WinupdatesV5.xml: Template to import into Zabbix.
- check_win_updates-v5.ps1: Powershell script. Must be placed in a "plugins" directory under your Zabbix Agent install folder (eg: C:\Program Files\Zabbix Agent\plugins).

## Installing

1. Ensure your Windows hosts have the agent installed and configured correctly and they are communicating with the Zabbix server.
2. Make sure you have edited the zabbix_agentd.conf file to include the line: EnableRemoteCommands=1 (<v5) or AllowKey=system.run[*] (>v5)
3. Download Winupdates.xml or WinupdatesV5.xml template from Github and import the template in Zabbix and apply it to one or more Windows Hosts.

### Option 1 - Deply Powershell Script Manually

Choose this option if you want to make any changes to the PowerShell script. You may want to change the report path, or you may need to change the agent install path if you don't have a default install (C:\Program Files\Zabbix Agent).

You also need to choose this option if you want to disable the auto-update feature.

1. Create a sub-folder called "plugins" in your Zabbix Agent install folder (eg: C:\Program Files\Zabbix Agent\plugins).
2. Download check_win_updates2020.ps1 (<v5) or check_win_updates-v5.ps1 (>v5) from Github and make any changes you need to it. Check the .ps file for instructions on making changes.
3. Copy check_win_updates2020.ps1 (<v5) or check_win_updates-v5.ps1 (>v5) to the plugins directory.

### Option 2 - Automatic Deployment of the Powershell script

This is a good option if you have a number of hosts you want to monitor but don't have an easy way to deploy the PS script to them all.

1. Edit the template in Zabbix (Configuration / Templates / WinUpdates).
2. Go to the "Items" tab.
3. You will see an item called "Deploy PS Script". Click on it to edit it.
4. Click the button to enable (if not ticked already) the item and click "Update".
5. Once all hosts have received the script, disable this item again, or it will re-download the script every day.

The item has a 1d update interval, so it may take up to a day for the PowerShell script to download. You can shorten this if you like.

### Dashboard

You can add a widget to your Dashboard by choosing type: Data overview and application "Winupdates-Panel".
