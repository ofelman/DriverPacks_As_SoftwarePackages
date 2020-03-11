<#
    Modify the entries to suite local requirements

    Dan Felman/HP Inc

    2/28/2020
        added $UseCM_OSFolders = $true                flag
#>

$OS = "Win10"
$OSVER = "1903"
$OSVALID = @("1809", "1903", "1909")

#-------------------------------------------------------------------
# Example systems to manage
#
$HPModelsTable = @(
	@{ ProdCode = '8438'; Model = "HP EliteBook x360 1030 G3" }
	@{ ProdCode = '83B3'; Model = "HP ELITEBOOK 830 G5" }
	@{ ProdCode = '844F'; Model = "HP ZBook Studio x360 G5" }
	@{ ProdCode = '83B2'; Model = "HP ELITEBOOK 840 G5" }
	@{ ProdCode = '8549'; Model = "HP ELITEBOOK 840 G6" }
	@{ ProdCode = '8470'; Model = "HP ELITEBOOK X360 1040 G5" }
    )

# this line assumes the script is running on a system with SCCM installed, so its PS modules are loaded
$CMSiteProvider = Get-PSDrive -PSProvider CMSite                                  # assume CM PS modules loaded at this time
$FileServerName = $CMSiteProvider.Root               

$DownloadSoftpaqDir = "\\$($FileServerName)\share\softpaqs"                       # + $OSVER
$ExtractPackageShareDir = "\\$($FileServerName)\share\packages\HP-"               # + $OSVER

$UseCM_OSFolders = $False
$LogFile = "$PSScriptRoot\HPDriverPackDownload.log"                               # Default to location where script runs from
