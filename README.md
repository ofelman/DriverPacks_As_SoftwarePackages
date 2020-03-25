# DriverPacks_As_SoftwarePackages
Script to automate creation of HP Driver Packages as Software in SCCM/SECM 

Based on a script created by GaryTown.com, this script adds a GUI and the ability to create the required Task Sequences used by the script
The following options are available:
  Report: describes available HP DriverPacks for the specific Windows 10 version, and what, if any, entries are available in the SCCM/SECM envirnment
  Download: if a new HP driverpack is available, it downloads it, unpacks it to the CM share, and updates CM as required
  Setup CM Task Sequences: If not yet in CM, it creates the 2 Task Sequences required by the script

Version 2.30 - checking the 'Use OS version Folders' will enable the Downloader script to move the driverpack packages in CM to folders named by the OS version; while unchecking it moves them to the Packages root folder

Version 2.31 - FIX: entries not checked are cleared when doing a Report or Download
Version 2.32
            - Fix: ini file proper update status of Use CM FOlders
Version 2.33
            - Fix: error creating CM Package if it didn't exist alreay...
                due to incorrect location for 'set-location C:' in function Option_Download
            - Fix: error listing info when Driverpack downloads successfully but CM Package not yet created
Version 2.35
            - Additional DriverPack download Fixes and cleanup - makes sure DriverPack > 350MB and extracted folder > 500MB
            - FIX - TS Step creation failed due to name (that included Model name) was > 50 chars - max allowed by CM TS Step
