# Requirements

This script requires the Virtual Deskop Powershell module to run. Install it from here: https://www.powershellgallery.com/packages/VirtualDesktop/ before running the script.

# Installation

This script is designed to be run via Windows Task Scheduler. Use these settings for your scheduled task:

Command: `"C:\Users\<username>\AppData\Local\Microsoft\WindowsApps\pwsh.exe"`

Arguments: `-ExecutionPolicy RemoteSigned -command & 'C:\Users\<username>\AppData\Local\Programs\BingWallpaper\Bingwallpaper.ps1'`
