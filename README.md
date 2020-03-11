# DriverPacks_As_SoftwarePackages
Script to automate creation of HP Driver Packages as Software in SCCM/SECM 

Based on a script created by GaryTown.com, this script adds a GUI and the ability to create the required Task Sequences used by the script
The following options are available:
  Report: describes available HP DriverPacks for the specific Windows 10 version, and what, if any, entries are available in the SCCM/SECM envirnment
  Download: if a new HP driverpack is available, it downloads it, unpacks it to the CM share, and updates CM as required
  Setup CM Task Sequences: If not yet in CM, it creates the 2 Task Sequences required by the script

Version 2.30 - checking the 'Use OS version Folders' will enable the Downloader script to move the driverpack packages in CM to folders named by the OS version; while unchecking it moves them to the Packages root folder

Version 2.31 - FIX: entries not checked are cleared when doing a Report or Download
