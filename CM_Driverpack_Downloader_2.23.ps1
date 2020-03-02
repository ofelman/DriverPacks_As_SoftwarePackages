<#  Creator @gwblok - GARYTOWN.COM (original script)
    Used to download Drivers Updates from HP
    This Script was created to build a Drivers Update Package. 
 
 
    REQUIREMENTS:  HP Client Management Script Library
    Download / Installer: https://ftp.hp.com/pub/caps-softpaq/cmit/hp-cmsl.html  
    Docs: https://developers.hp.com/hp-client-management/doc/client-management-script-library
 
    Modes... Report - only compares HP against what is in SCCM (No downloading)
             Download - Downloads DriverPacks and Updates SCCM as needed
   
    Dan Felman/HP Inc: Reformatted script to support Report, Download properly
    Changes:
        - when a CM Package is not created or available in CM, the script will now
            create a Package and show it on the Gui... 
        - removed the need to maintain CM Package IDs... it is all resolved in the code  
        - option Download will now create a CM S/W Package if none exists    
        - Added ability to report on CM Task Sequence Module and Steps
        - Added ability to create NEW steps in TS Module as Product/Model entries are added to script 
        - added ability to select what models to Report or Download during runtime
        - added ability to change OS versions (supported versions listed in ini.ps1 file)
        - some color coding in Data view to show if checking a OS version but CM has a different version in the S/W package
        - GUI Front End
        # DriverPack information publised in PS DataGrid table
        Ver 2.2 - replaced WIndows OS version entry with combobox of allowed versions from PS1 file
            - moved Debug Output checkmark near the output box
            - added OS version to Softpaq version on CM's package Version header (e.g. '1903 - 10.0')
        Ver 2.23
            - added Tooltips to Gui fields
            - modified function Update_SoftpaqFolders to clean up the Temp folder correctly, where there
                were some ACL permissions issues... If script is run with Admin rights, it should be fine
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $false,Position = 1,HelpMessage = "Application")]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("Report","Download")]
	$RunMethod = "Report"
)
$ScriptVersion = "2.23 Gui (2/18/2020)"

$scriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

#Settings for Google - Rest of Info @ Bottom
#$SMTPServer = "smtp.gmail.com"
#$SMTPPort = "587"
#$Username = "username@gmail.com"
#$Password = "password"

#Script Vars
#$EmailArray = @()

<#
    These settings are needed to support this Drivers as Software Packages script
#>

# This Task Sequence will hold ALL Driver Package steps, and is required
$TSDriverModuleName = "Driver Package Module"
$TSDriverPackageModule_Desc = "This task hosts the downloaded package content steps used for driver injection"

# The Driver Injection Module is called by the Deployment Task Sequence and will not be modified once set up
$TSDriverInjectionName = "Driver Injection Module"
$TSDriverInjectionModule_Desc = "This task calls the Driver Package Module to download the driver package to the local disk and then uses DISM to inject drivers"

#--------------------------------------------------------------------------------------
#Reset Vars

$AdminRights = $false
$TSDriverModule = $null                            # This Task Sequence will hold ALL Driver Package steps
$CMConnected = $false
$IniDataset = $null

$Debug_Output = $false
$ErrorType = 3
$TypeDebug = 4
$SuccessType = 5
$TypeNoNewline = 10

#--------------------------------------------------------------------------------------
# check for ConfigMan PS module on this server

#$CMInstall = "E:\Program Files\Microsoft Configuration Manager\AdminConsole\bin" # NOW in Gui

if (Test-Path $env:SMS_ADMIN_UI_PATH) {
    $CMInstall = Split-Path $env:SMS_ADMIN_UI_PATH
    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
} else {
    Write-Host "Can't find CM Installation on this system"
    exit
}

#--------------------------------------------------------------------------------------
#Script Vars Environment Specific loaded from INI.ps1 file

$IniFile = "CM_Driverpack_ini.ps1"                     # version 2.0 remove CM Package ID from product list
$IniFIlePath = "$($ScriptPath)\$($IniFile)"

# source the code in the INI file

. $IniFIlePath

$SiteCode = $CMSiteProvider.Name                     

$IniFileLines = (Get-Content -Path $IniFIlePath).Length                 # !!!!!!!! for future use - Length = lines in the file !!!!!!!!


#=====================================================================================
# Check for Admin rights
$oIdentity= [Security.Principal.WindowsIdentity]::GetCurrent()
$oPrincipal = New-Object Security.Principal.WindowsPrincipal($oIdentity)
if($oPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator )){
    Write-Warning "Please start script with Administrator rights! Exit script"
    $AdminRights = $true
}

#=====================================================================================
#region: CMTraceLog Function formats logging in CMTrace style
function CMTraceLog {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false)] $Message,
		[Parameter(Mandatory = $false)] $ErrorMessage,
		[Parameter(Mandatory = $false)] $Component = "HP DriverPack Downloader",
		[Parameter(Mandatory = $false)] [int]$Type,
		[Parameter(Mandatory = $true)] $LogFile
	)
	<#
    Type: 1 = Normal, 2 = Warning (yellow), 3 = Error (red)
    #>
	$Time = Get-Date -Format "HH:mm:ss.ffffff"
	$Date = Get-Date -Format "MM-dd-yyyy"

	if ($ErrorMessage -ne $null) { $Type = $ErrorType }
	if ($Component -eq $null) { $Component = " " }
	if ($Type -eq $null) { $Type = 1 }

	$LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"`" file=`"`">"

    #$Type = 4: Debug output = $TypeDebug
    #$Type = 10: no \newline

    if ( ($Type -ne $Script:TypeDebug) -or ( ($Type -eq $Script:TypeDebug) -and $Debug_Output) ) {
        $LogMessage | Out-File -Append -Encoding UTF8 -FilePath $LogFile
        OutToForm $Message $Type
        
    } else {
        $lineNum = ((get-pscallstack)[0].Location -split " line ")[1]    # output: CM_Driverpack_Downloader.ps1: line 557
        #Write-Host "$lineNum $(Get-Date -Format "HH:mm:ss") - $Message"
    }

} # function CMTraceLog

#=====================================================================================

Function OutToForm { 
	[CmdletBinding()]
	param( $pMessage, [int]$pmsgType )

    switch ( $pmsgType )
    {
        1 { 
            # default color is black
          }
        2 { 
            $TextBox.SelectionColor = "Brown"
          }
        3 { 
            $TextBox.SelectionColor = "Red"                   # Error
          }
        4 { 
            $TextBox.SelectionColor = "Orange"                # Debug Output
          }
        5 { 
            $TextBox.SelectionColor = "Green"                 # success details
          }
        10 { 
            # do NOT add \newline to message output
          }

    } # switch ( $pmsgType )

    if ( $pmsgType -eq $Script:TypeDebug ) {
        $pMessage = "<dbg>"+$pMessage
    }

    if ( $pmsgType -eq $TypeNoNewline ) {
        $TextBox.AppendText($pMessage+"")
    } else {
        $TextBox.AppendText($pMessage+"`n")
    }
    $TextBox.Refresh()
    $TextBox.ScrollToCaret()

} # Function OutToForm

Function List_CMObkects {
    Import-Module $env:SMS_ADMIN_UI_PATH.Replace("\bin\i386","\bin\configurationmanager.psd1")
    Set-Location -Path "$($SiteCode):" | Out-Null
 
    $sitecode = $CMSiteProvider
    $colltomove = "Ella1W10"
    $destcollfolder = "$($sitecode):\Packages" 
 
    $collID = Get-CMCollection -Name $colltomove
    Move-CMObject -InputObject $collID -FolderPath $destcollfolder
}

#=====================================================================================
<#
    Function Test_CMConnection
        The function will test the CM server connection
        and that the Task Sequences required for use of the Script are available in CM
        - will also test that both download and share paths exist
#>
Function Test_CMConnection {

    $boolConnectionRet = $False

    CMTraceLog -Message "Connecting to CM Server: ""$FileServerName""" -Type $TypeNoNewline -LogFile $LogFile
    
    if (Test-Path $CMInstall) {
        
        try { Test-Connection -ComputerName "$FileServerName" -Quiet

            CMTraceLog -Message " ...Connected" -Type $SuccessType -LogFile $LogFile
            $boolConnectionRet = $True
            #$EmailArray += "<font color=Black>Script was run in Mode: $($RunMethod)</font><br>"
            #$EmailArray += "<font color=Black>Connecting to File Share Server: $($FileServerName)</font><br>"

            # -----------------------------------------------------------            
            # Now, let's make sure the CM Drivers Module Task Sequence is in place

            CMTraceLog -Message "Testing for Task Sequence: ""$($TSDriverModuleName)""" -Type $TypeNoNewline -LogFile $LogFile

            Set-Location -Path "$($SiteCode):"
            $Script:TSDriverModule = Get-CMTaskSequence -Name $TSDriverModuleName
            if ($TSDriverModule) {
                CMTraceLog -Message " ...Task Sequence ($($TSDriverModule.PackageID)) Found" -Type $SuccessType -LogFile $LogFile
            } else {
                $boolConnectionRet = $False
		        CMTraceLog -Message " ...Task Sequence NOT Found - REQUIRED" -Type $ErrorType -LogFile $LogFile
	        } # else
	        Set-Location -Path "C:"
            # -----------------------------------------------------------
        }
        catch {
	        CMTraceLog -Message "Not Connected to File Server, Exiting" -Type $ErrorType -LogFile $LogFile
        }
    } else {
        CMTraceLog -Message "CM Installation path NOT FOUND: '$CMInstall'" -Type $ErrorType -LogFile $LogFile
    } # else
    
    #================================================================================
    # Let's now also test for the Download and Share paths 
    #
    $DownloadSoftpaqDirToCheck = $DownloadSoftpaqDir.Replace('\'+$OSVER,"")
    CMTraceLog -Message "Testing for Download Path: ""$($DownloadSoftpaqDirToCheck)""" -Type $TypeNoNewline -LogFile $LogFile
    if (Test-Path -Path $DownloadSoftpaqDirToCheck -IsValid) {
        CMTraceLog -Message " ...Path is valid" -Type $SuccessType -LogFile $LogFile
    } else {
        CMTraceLog -Message " ...Path is NOT valid" -Type $ErrorType -LogFile $LogFile
        $boolConnectionRet = $false
    }
    
    $ExtractPackageShareDirToCheck = $ExtractPackageShareDir.Replace('\HP-'+$OSVER,"")
    CMTraceLog -Message "Testing for Share Path: ""$($ExtractPackageShareDirToCheck)""" -Type $TypeNoNewline -LogFile $LogFile
    if (Test-Path -Path $ExtractPackageShareDirToCheck -IsValid) {
        CMTraceLog -Message " ...Path is valid" -Type $SuccessType -LogFile $LogFile
    } else {
        CMTraceLog -Message " ...Path is NOT valid" -Type $ErrorType -LogFile $LogFile
        $boolConnectionRet = $false
    }

    return $boolConnectionRet

} # Function Test_CMConnection

#=====================================================================================
<#
    Function CM_CheckTS
    check on, or create the 2 Task Sequences used by the script packages
#>

Function CM_CheckTS { 

    CMTraceLog -Message "[CM_CheckTS()] enter" -Type $TypeDebug -LogFile $LogFile        # Debug_Output=$true

    CMTraceLog -Message "Checking CM for Driver Module and Injection Task Sequences" -Type 1 -LogFile $LogFile

	Set-Location -Path "$($SiteCode):"

    #######################################################################################################
    # Driver Package Module

    $TSDriverModule = Get-CMTaskSequence -Name $TSDriverModuleName

    if ( $TSDriverModule ) {
        CMTraceLog -Message "... Task Sequence: ""$($TSDriverModule.Name)"" Found" -Type 1 -LogFile $LogFile
    } else {
        CMTraceLog -Message "... Task Sequence: ""$($TSDriverModuleName)"" NOT Found" -Type 2 -LogFile $LogFile
        # let's create the Task Sequence ir asked for
        CMTraceLog -Message "... Adding TS to SCCM" -Type $TypeDebug -LogFile $LogFile
        $TSDriverModule = New-CMTaskSequence -CustomTaskSequence -Name $TSDriverModuleName -Description $TSDriverPackageModule_Desc
        # create the new group
        $PackageGroup = New-CMTaskSequenceGroup -Name "Driver Packages as Software Packages"
        Add-CMTaskSequenceStep -InsertStepStartIndex 0 -TaskSequenceName $TSDriverModule.Name -Step $PackageGroup
        CMTraceLog -Message "... TS ""$($TSDriverModule.Name)"" created." -Type 1 -LogFile $LogFile
    } # else

    #######################################################################################################
    # Driver Injection Module - called by Deployment Task Sequence

    $TSInjectionModule = Get-CMTaskSequence -Name $TSDriverInjectionName

    if ( $TSInjectionModule ) {
        CMTraceLog -Message ".... Task Sequence: ""$($TSInjectionModule.Name)"" Found" -Type 1 -LogFile $LogFile
    } else {
        CMTraceLog -Message ".... Task Sequence: ""$($TSDriverInjectionName)"" NOT Found" -Type 2 -LogFile $LogFile
        # let's create the Task Sequence ir asked for
        CMTraceLog -Message "... Adding TS to SCCM" -Type $TypeDebug -LogFile $LogFile
        $TSInjectionModule = New-CMTaskSequence -CustomTaskSequence -Name $TSDriverInjectionName -Description $TSDriverInjectionModule_Desc
        # create the new group
        $InjectionGroup = New-CMTaskSequenceGroup -Name "Driver Management"
        Add-CMTaskSequenceStep -InsertStepStartIndex 0 -TaskSequenceName $TSInjectionModule.Name -Step $InjectionGroup
        
        # add the Run Task Sequence Module step
        $TSDriverModuleSubStep = New-CMTSStepRunTaskSequence -Name $TSDriverModule.Name -RunTaskSequence $TSDriverModule
        Set-CMTaskSequenceGroup -InputObject $TSInjectionModule -StepName $InjectionGroup.Name -AddStep $TSDriverModuleSubStep -InsertStepStartIndex 0

        # add the Run DISM step w/option to run ONLY if the copy succeeded (e.g. %_SMSTSMDataPath%\Drivers exists)
        
        $DISMCommand = "DISM.exe /Image:%OSDTargetSystemDrive%\ /Add-Driver /Driver:%_SMSTSMDataPath%\Drivers\ /Recurse /logpath:%_SMSTSLogPath%\dism.log"
        $DISMCommand_Desc = "Run DISM to inject all drivers downloaded to the client - Use /Recurse option"

        $lQuery = New-CMTSStepConditionFolder -FolderPath "%_SMSTSMDataPath%\Drivers"
        $IfAny = New-CMTSStepConditionIfStatement -StatementType Any -Condition $lQuery
        $DISMCmdSubStep = New-CMTSStepRunCommandLine -CommandLine $DISMCommand -Name "Install Downloaded Drivers via DISM /Recurse" -Description $DISMCommand_Desc -Condition $IfAny
        Set-CMTaskSequenceGroup -InputObject $TSInjectionModule -StepName $InjectionGroup.Name -AddStep $DISMCmdSubStep -InsertStepStartIndex 1

        CMTraceLog -Message "... TS ""$($TSInjectionModule.Name)"" created." -Type 1 -LogFile $LogFile

    } # else
	Set-Location -Path "C:"
    CMTraceLog -Message "Done Updating Driver Module and Injection Task Sequences" -Type 1 -LogFile $LogFile
    CMTraceLog -Message "[CM_CheckTS()] exit" -Type $TypeDebug -LogFile $LogFile        # Debug_Output=$true

} # Function CM_CheckTS

#=====================================================================================
<#
    Function Download_Softpaq
        Using a HP CMSL cmdlet, it downloads an HP DriverPack softpaq
#>
Function Download_Softpaq { 
	[CmdletBinding()]
	param( $pHPDriverPackID,[string]$pPath )

    Get-Softpaq -number $pHPDriverPackID -saveAs $pPath -overwrite yes -Extract     # -DestinationPath "$($env:temp)\SPExtract\$($HPModel.Model)" 
    Get-SoftpaqMetadataFile $pHPDriverPack.id  -SaveAs $pPath -Overwrite "yes"      # get the latest CVA file available (e.g. 'sp99673')

    # let's compare hashes - MD5 from cva and downloaded files
    # for future use - to make sure the donwnloaded Softpaq matches the original
    $SoftpaqMD5 = Get-SoftpaqMetadata $pHPDriverPackID | Out-SoftpaqField SoftPaqMD5
    $ComputedMD5 = (Get-FileHash $pPath -Algorithm MD5)
    CMTraceLog -Message "$($pPath) - cva hash=$($SoftpaqMD5), PS hash=$($ComputedMD5)" -Type $TypeDebug -LogFile $LogFile        # Debug_Output=$true

    CMTraceLog -Message "... Softpaq '$($pPath)' downloaded" -Type 1 -LogFile $LogFile    

} # Function Download_Softpaq

#=====================================================================================
<#
    Function Update_SoftpaqFolders
    downloads new driver pack softpaq and extracts it for CM use
    returns Path of extracted folder in CM share for use in CM Drive Package
    Parameters:
        HP DriverPack - PS Object
        HP Model Name - Product to check on
#>
function Update_SoftpaqFolders {
	[CmdletBinding()]
	param($pHPDriverPack,[string]$pModelName)

    $Downloaded = $false                                # Start where nothing is downloaded
	[bool]$CMPackageNeedsUpdate = $false

    # ... Reset the paths for the DriverPack and the extracted OS versions we are searching for...
    # include OSVER version info in Softpaq and Packages folders

    # (ex: Softpaq source - \\...\share\softpaqs\1909\HP EliteBook x360 1030 G3\9.00)
	$Script:DownloadDriverPackFullPath = "$($DownloadSoftpaqDir)\$($Script:OSVER)\$($pModelName)\$($pHPDriverPack.Version)"

    # (ex: Unpack source - \\...\share\packages\HP-1909\HP EliteBook x360 1030 G3)
	$script:ExtractPackageShareDirCurrent = "$($ExtractPackageShareDir)$($Script:OSVER)\$($pModelName)" # this is the folder in the share that SCCM will draw from
   
    $DriverPackFullPathExe = $DownloadDriverPackFullPath + "\$($pHPDriverPack.id).exe"     # name of full path to driverpack, including SP exe

    if (Test-Path $DownloadDriverPackFullPath) {                   # e.g. is the driverpack path already exist?
		if (Test-Path $DriverPackFullPathExe) {                    #         then, see if the SPXXXXXX.exe exists - at least a version of it
			CMTraceLog -Message "... Softpaq '''$($DriverPackFullPathExe)' exists" -Type 1 -LogFile $LogFile
		} else {
			CMTraceLog -Message "... Softpaq '$($DriverPackFullPathExe)' does NOT exist. Let's download the Softpaq" -Type 2 -LogFile $LogFile
            Download_Softpaq $pHPDriverPack.id $DriverPackFullPathExe
			$Downloaded = $true
		} # else

	} else {
		CMTraceLog -Message "... Softpaq path '$($DownloadDriverPackFullPath)' does NOT exist. Let's create it and download the softpaq" -Type 2 -LogFile $LogFile
		New-Item -Path $DownloadDriverPackFullPath -ItemType Directory -Force # create folder for Softpaq download
        Download_Softpaq $pHPDriverPack.id $DriverPackFullPathExe
		$Downloaded = $true     
    } # if (Test-Path $DownloadDriverPackFullPath) 

    # at this point, we have a folder and the DriverPack softpaq downloaded

	if (Test-Path $script:ExtractPackageShareDirCurrent) {
    <#
        let's work on the unpacking, if needed to use with a CM Software Package
        ALso, make sure the unpacked driver pack folder is current - use version info from XML file
        e.g. compare XML version info between downloaded DriverPack Softpaq and the extracted/already in place DriverPack

        ----------- HP DriverPack Manifest XML file sample / a PS Package contents ------------------
        <Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04">
            <Obj RefId="0">
                <TN RefId="0">
                    <T>Selected.System.Management.Automation.PSCustomObject</T>
                    <T>System.Management.Automation.PSCustomObject</T>
                    <T>System.Object</T>
                </TN>
                <MS>
                    <S N="Id">sp99708</S>
                    <S N="Name">HP Elite/ZBook 8x0 G5 Windows 10 x64 Driver Pack</S>
                    <S N="Category">Manageability - Driver Pack</S>
                    <S N="Version">10.0</S>
                    <S N="Vendor">HP</S>
                    <S N="ReleaseType">Routine</S>
                    <S N="SSM">false</S>
                    <S N="DPB">false</S>
                    <S N="Url">ftp.hp.com/pub/softpaq/sp99501-100000/sp99708.exe</S>
                    <S N="ReleaseNotes">ftp.hp.com/pub/softpaq/sp99501-100000/sp99708.html</S>
                    <S N="Metadata">ftp.hp.com/pub/softpaq/sp99501-100000/sp99708.cva</S>
                    <S N="MD5">0688611b9337d4e28685715de5b71aa9</S>
                    <S N="Size">848385496</S>
                    <S N="ReleaseDate">2019-11-22</S>
                </MS>
            </Obj>
        </Objs>
    #>

		######################################################################
        # let's see if a valid driverpack XML file exists and it is up to date
        ######################################################################
        $PackageXMLfile = "$($script:ExtractPackageShareDirCurrent)\DriverPackInfo.XML"
	    CMTraceLog -Message "... checking for CM Packge share xml file: '$($PackageXMLfile)'" -Type 1 -LogFile $LogFile

		if (Test-Path $PackageXMLfile) {

            # Let's look at the contents of the XML file
			$xmlFile = New-Object -TypeName XML
			$xmlFile.Load($PackageXMLfile)
			$xmlPackageCurrVersion = ($xmlFile.Objs.obj.MS.S | Where-Object { $_.N -eq 'Version' })."#text"

			if ($xmlPackageCurrVersion -eq $pHPDriverPack.Version) {
				CMTraceLog -Message "... CM package folder is up to date - (XML file versions match: $($xmlPackageCurrVersion))" -Type 1 -LogFile $LogFile
			} else {
				CMTraceLog -Message "... CM package folder is NOT up to date - (XML file versions do NOT match: $($xmlPackageCurrVersion) vs HP $($pHPDriverPack.Version ))" -Type $ErrorType -LogFile $LogFile
				$CMPackageNeedsUpdate = $true
			} # else

		} else { # if (Test-Path $PackageXMLfile)
			$CMPackageNeedsUpdate = $true                # XML file does not exist (e.g. drivepack needed)
		} # else

	} else {

		CMTraceLog -Message "CM Package folder does not exist, will create it" -Type 2 -LogFile $LogFile
		New-Item $script:ExtractPackageShareDirCurrent -ItemType Directory -Force
		CMTraceLog -Message "CM Package folder created" -Type 1 -LogFile $LogFile
		$CMPackageNeedsUpdate = $true

	} # else if (Test-Path $script:ExtractPackageShareDirCurrent) 

    ########################################################################
    # DriverPack is available... let's see if we need to update the share
    ########################################################################

    if ( $CMPackageNeedsUpdate ) {
        
        #------------------------------------------------------------------------------
        # Let's make sure the root Temp folder to work on is in place, if not create it
        $TempDir = "$($env:TMP)\SPExtract"
        if ( !(Test-Path $TempDir) ) {
            New-Item $TempDir -ItemType Directory -Force 
        } 
        $TempDir = (Get-Item -Literalpath ("$($env:TMP)\SPExtract")).fullname                # expand to avoid path truncation

        # now, let's remove any folders in our Temp and proper share folders
        CMTraceLog -Message "Removing Temp folders: $($TempDir)\*" -Type 1 -LogFile $LogFile  
        #Remove-Item -Path "$($TempDir)\*" -Recurse -Force -verb runAs -ErrorAction SilentlyContinue      # need Admin??? rights for this
        
        #powershell -NoLogo -NonInteractive -File C:\Scripts\Backup.ps1 -Param1 TestBackup

        #Start-Process powershell.exe -Verb "runAs" -ArgumentList "-Command &{Remove-Item $TempDir\* -Recurse -Force}" -windowstyle "Minimized"
        Start-Process powershell.exe -verb "runAs" -ArgumentList "-Command &{Remove-Item $TempDir\* -Recurse -Force}" -windowstyle "Minimized"

        $ModelTempDir = "$($TempDir)\$($pModelName)"                                         # add Folder name of product      
        
        CMTraceLog -Message "Recreating $($ModelTempDir)" -Type 1 -LogFile $LogFile
        New-Item $ModelTempDir -ItemType Directory -Force 

		CMTraceLog -Message "Removing Model share path:$($ExtractPackageShareDirCurrent)" -Type 1 -LogFile $LogFile
		Remove-Item -Path $ExtractPackageShareDirCurrent -Recurse -Force -ErrorAction SilentlyContinue
        
        #------------------------------------------------------------------------------
		# now, populate the Temp folder... extract contents of DriverPack Softpaq... 
        CMTraceLog -Message "Unpacking $($DriverPackFullPathExe) to $($ModelTempDir)" -Type 1 -LogFile $LogFile
        
        Start-Process -Verb RunAs $DriverPackFullPathExe -ArgumentList "-e -s -f""$($ModelTempDir)""" -Wait

        # find the created subfolder
        $SubDir = Get-ChildItem -Path $ModelTempDir -Attributes "Directory" | Sort-Object LastAccessTime -Descending | Select-Object -First 1                                         
        $newCopyPath = "$($ModelTempDir)\$($SubDir)"                                        # (e.g. full path ...\EB_X360_1030_G3)
        
        # ... TBD ...NEXT: should check for OS folder that matches OSVER we are searching for... TBD ...
        # (e.g. wt64_1809, or wt64_1903, etc. - there may be one or multiple folders potentially)

        $CopyWinFolder = Get-ChildItem -Path "$newCopyPath" -Attributes "Directory" | Sort-Object LastAccessTime -Descending | Select-Object -First 1       
        $CopyFromDir = "$($newCopyPath)\$($CopyWinFolder)"                                  # (e.g. full path ...\EB_X360_1030_G3\wt64_1903)

        #------------------------------------------------------------------------------
        # Next get the unpacked drivers to the SCCM share corresponding to the Model and OS version
		CMTraceLog -Message "Copying extracted contents to $($ExtractPackageShareDirCurrent)" -Type 1 -LogFile $LogFile
        Copy-Item -path $CopyFromDir -Destination $ExtractPackageShareDirCurrent -Force -Recurse

        # create an XML file for the DriverPack on the share directory
        CMTraceLog -Message "Creating  DriverPack XML file" -Type 1 -LogFile $LogFile
		Export-Clixml -InputObject $pHPDriverPack -Path "$($ExtractPackageShareDirCurrent)\DriverPackInfo.XML"

		#------------------------------------------------------------------------------
        # clean up Temp folder used - fails without SysAdmin privileges !!!
		if (Test-Path $TempDir) {
			CMTraceLog -Message "Removing Temp folder $($TempDir)" -Type 1 -LogFile $LogFile
		    Start-Process powershell.exe -verb "runAs" -ArgumentList "-Command &{Remove-Item $TempDir\* -Recurse -Force}" -windowstyle "Minimized"
		} # if (Test-Path $TempDir)

    } # if ($CMPackageNeedsUpdate) 

} # function Update_SoftpaqFolders

#=====================================================================================
<#
    function Set_HPDriverPackage
        updates a CM Driver Package and updates CM's Distribution Points

    REQUIREMENTS:
        Driver Package
        DriverPack Name/ID (e.g. sp99111)
        DriverPack Version #
        DriverPack Path on CM Share
    RETURNS:
        updated CM Software Package
#>
function Set_HPDriverPackage {
	[CmdletBinding()]
	param( $psPkg, $psSoftpaqID, $psPkgVersion, $psPath )
    
    CMTraceLog -Message "[Set_HPDriverPackage()] enter" -Type $TypeDebug -LogFile $LogFile        # Debug_Output=$true
    CMTraceLog -Message "..> psPkgName=$($psPkg.name), psPkgID=$($psSoftpaqID), psPkgVersion=$($psPkgVersion), psPath=$($psPath) <<<" -Type $TypeDebug -LogFile $LogFile        # Debug_Output=$true

    Set-Location -Path "$($SiteCode):"

	Set-CMPackage -Name $psPkg.Name -Language $psSoftpaqID
	Set-CMPackage -Name $psPkg.Name -Version "$($OSVER) - $($psPkgVersion)"
	Set-CMPackage -Name $psPkg.Name -Path $psPath

	update-CMDistributionPoint -PackageId $psPkg.PackageID

	Set-Location -Path "C:"

    CMTraceLog -Message "[Set_HPDriverPackage()] exit" -Type $TypeDebug -LogFile $LogFile        # Debug_Output=$true

    Return $psPkg

} # function Set_HPDriverPackage

#=====================================================================================
<#
    Function Update_TSPackageStep
    expects parameter "Driver Package Module TS", "PackageId" , "Package Name"
    Returns String: - if Step found and Associated Package was correct
                    - if Step found but Package ID was incorrect and was corrected
                    - if Step was NOT found and Package Step was added to the Package Module TS
#>
function Update_TSPackageStep {
	[CmdletBinding()]
	param($pTSDriverPackageModule,$pTSDriverPackage,$pProductCode, $pReportOnly)

    CMTraceLog -Message "[Update_TSPackageStep()] enter" -Type $TypeDebug -LogFile $LogFile        # Debug_Output=$true
    CMTraceLog -Message "..> pProductCode=$($pProductCode); pReportOnly=$($pReportOnly) <<<" -Type $TypeDebug -LogFile $LogFile        # Debug_Output=$true

	$lSTEPFoundAndUpToDate = "Download Package STEP was Found in TS:""$($TSDriverModule.Name)"" and Matches CM Package"
	$lSTEPFoundButUpdated = "Download Package STEP was Found, but was Updated"
	$lSTEPNotFoundCreated = "Download Package STEP was Not Found. It was Created"
	$lSTEPFoundButNeedsUpdated = "Download Package STEP was Found, but Package ID is incorrect and Needs Updated"
	$lSTEPNotFound = "Download Package STEP was NOT Found"

	$lStepReturn = $lSTEPFoundAndUpToDate # Assume Package STEP is OK
    [bool]$lStepFound = $false
	$lTSStep = $null

	Set-Location -Path "$($SiteCode):"

	if ($pTSDriverPackageModule -ne $null) {

		# Find all steps already in the Module
		$TSDriverModuleSteps = ($pTSDriverPackageModule | CMTaskSequenceStep)

		# See if the PackageId is already in the Module TS
		$lPosition = 0
		foreach ($TSDriverStep in $TSDriverModuleSteps) {
			if (($TSDriverStep.DownloadPackages -eq $pTSDriverPackage.PackageID) -or ($TSDriverStep.PackageInfo.Name -eq $pTSDriverPackage.Name)) {
				$lPosition += 1
				$lTSStep = $TSDriverStep
                $lStepFound = $true
				continue
			} # if ()
		} # foreach ()

		<#
            If Download Package Step is in the Module, make sure it referrences the correct Package ID
            If Package Step is NOT found, then add it
        #>

		# but, first... Build out the -Condition with the appropriate WMI Query - JUST IN CASE We Need to Add or Mod the Step
		$sWMIQuery = "select * from win32_baseboard where product = ""$($pProductCode)"" "
		$lQuery = New-CMTSStepConditionQueryWmi -Query $sWMIQuery
		$IfAny = New-CMTSStepConditionIfStatement -StatementType Any -Condition $lQuery

		if ( $lStepFound ) {

			# $TSDriverStep was found

			if ($lTSStep.DownloadPackages -ne $pTSDriverPackage.PackageID) {

				if ($pReportOnly) {

					$lStepReturn = $lSTEPFoundButNeedsUpdated

				} else {

					# need to update the Package info in the Step
					# ... Since we can't just update the PackageID in the STEP, Recreate it w/the correct PackageID

					Remove-CMTaskSequenceStep -InputObject $pTSDriverPackageModule -StepName $pTSDriverPackage.Name -Force

					# Create the CM TS Download Software Pckage Step
					$lTSDriverPackageStep = New-CMTaskSequenceStepDownloadPackageContent -AddPackage $pTSDriverPackage -Name "Download $($pTSDriverPackage.Name) Drivers" -LocationOption CustomPath -Path "%_SMSTSMDataPath%\Drivers" -Description "New TS created with PowerShell" -Condition $IfAny
					# ... and add it to the correct TS Module
					$lAddedPackageStep = Set-CMTaskSequenceGroup -InputObject $pTSDriverPackageModule -AddStep $lTSDriverPackageStep -InsertStepStartIndex $lPosition
					update-CMDistributionPoint -PackageId $pTSDriverPackage.PackageID
					$lStepReturn = $lSTEPFoundButUpdated

				} # else

			} # if ( $TSDriverStep.DownloadPackages -ne $pTSDriverPackage.PackageId )

		} else {     # step in TS NOT Found

			if ($pReportOnly) {
				$lStepReturn = $lSTEPNotFound
                $lStepFound = $false
			} else {

				# Create the CM TS Step   
				$lTSDriverPackageStep = New-CMTaskSequenceStepDownloadPackageContent -AddPackage $pTSDriverPackage -Name "Download $($pTSDriverPackage.Name) Drivers" -LocationOption CustomPath -Path "%_SMSTSMDataPath%\Drivers" -Description "New TS created with PowerShell" -Condition $IfAny
				# add it to the correct TS Module
				$lAddedPackageStep = Set-CMTaskSequenceGroup -InputObject $pTSDriverPackageModule -AddStep $lTSDriverPackageStep -InsertStepStartIndex $lPosition

				update-CMDistributionPoint -PackageId $pTSDriverPackage.PackageID

				$lStepReturn = $lSTEPNotFoundCreated
                $lStepFound = $true

			} # if ( $pReportOnly )

		} # else

	} # ( $pTSDriverPackageModule -ne $null )

	Set-Location -Path "C:"

    CMTraceLog -Message "... $($lStepReturn): ""$($pTSDriverPackage.Name)""\$($pTSDriverPackage.PackageID)->$($pTSDriverPackage.version)" -Type 1 -LogFile $LogFile

    CMTraceLog -Message "[Update_TSPackageStep()] exit" -Type $TypeDebug -LogFile $LogFile        # Debug_Output=$true

	return $lStepFound

} # Update_TSPackageStep

#=====================================================================================
<#
    Function ListView_FillInItem
        reports items/row on the Model List, like
            Softpaq Name/ID - column 4
            Softpaq Version - column 5
#>
Function ListView_FillInItem {
    [CmdletBinding()]
	param($pModelsList, $pRow, $pHPDriverpack, $pTSStepFound)

    CMTraceLog -Message "ListView_FillInItem enter" -Type $Script:TypeDebug -LogFile $LogFile 

    # get the model we are looking for
    $lEntryModel = $pModelsList[2,$pRow].Value
    # and find the CM Package to report on
    Set-Location -Path "$($SiteCode):"
    $lCMDriverPackage = Get-CMPackage -Name $lEntryModel -Fast

    Set-Location -Path "C:"

    # items are rows, subitems are columns
    # Write-Host -ForegroundColor Yellow $pRow $pModelsList.items[$pRow].subitems[2]
    if ( $pHPDriverpack -eq $null ) {
        $pModelsList[3,$pRow].clear                                    # Available
        $pModelsList[4,$pRow].clear                                    # CM Ver
        $pModelsList[5,$pRow].clear                                    # CM OS
        $pModelsList[6,$pRow].clear                                    # CM Package
        $pModelsList[7,$pRow].clear                                    # Step
    } else {
        $OSVersionStart = $lCMDriverPackage.PkgSourcePath.IndexOf("HP-")                          #     find the package source path's "HP-1909" string location
        $OSVersion = $lCMDriverPackage.PkgSourcePath.Substring(($OSVersionStart+3),4)             #     Get the Windows Version (e.g. 1903, 1909, etc)
        
        $pModelsList[3,$pRow].Value = ($pHPDriverpack.Id+' - '+$pHPDriverpack.Version)            # Softpaq#
        
        $lPkgVersion = $lCMDriverPackage.Version.split(" ")[-1]#"one - three".split(" ")[-1]      # Get the Softpaq Version in the CM Package (e.g. from '1903 - 8.0' pick off 8.0
        $pModelsList[4,$pRow].Value = $lPkgVersion                                                #$lCMDriverPackage.Version                                   # Softpaq Version 
        if ( $pHPDriverpack.Version -ne $lPkgVersion ) {
            $pModelsList[4,$pRow].Style.ForeColor = "Red"
        } else {
            $pModelsList[4,$pRow].Style.ForeColor = "Black"
        } # if
        
        $pModelsList[5,$pRow].Value = $OSVersion                                                  # OS Version    
        if ( $OSVersion -ne $OSVER ) {
            $pModelsList[5,$pRow].Style.ForeColor = "Brown"
        } else {
            $pModelsList[5,$pRow].Style.ForeColor = "Black"
        } # if else
                
        $pModelsList[6,$pRow].Value = $lCMDriverPackage.PackageId                                 # Actual CM Package ID
        
        if ( $pTSStepFound ) {
            $pModelsList[7,$pRow].Value = 'v'                                                     # Package Step exists in TS?
        } else {
            $pModelsList[7,$pRow].Clear
        }
    } # else
    $pModelsList.refresh
    CMTraceLog -Message "ListView_FillInItem exit: Pkg Softpaq Ver: $($lCMDriverPackage.Version); Pkg OS Version $($OSVersion)" -Type $Script:TypeDebug -LogFile $LogFile 

} # Function ListView_FillInItem

#=====================================================================================
<#
#>
Function CM_Report {
    [CmdletBinding()]
	param( $pModelsList, $pCheckedItemsList)

    CMTraceLog -Message "Reporting" -Type 1 -LogFile $LogFile

    CMTraceLog -Message "OS Version to check for: $($Script:OSVER)" -Type $Script:TypeDebug -LogFile $LogFile 

    # run thru every checked row # in the Models listView box
    $pCheckedItemsList | ForEach-Object  {

        # SubItems[0] is the checkmark column
        $lEntryProdCode = $pModelsList[1,$_].Value
        $lEntryModel = $pModelsList[2,$_].Value

        #write-host "$($lEntryProdCode) - $($lEntryModel)"

        # additional debug output, if enabled
        CMTraceLog -Message "Checking: row:$($_); Model=/$($lEntryModel)/; ProdCod=/$($lEntryProdCode)" -Type $TypeDebug -LogFile $LogFile   

        CMTraceLog -Message "(Checking: ProdCode:'$($lEntryProdCode)', '$($lEntryModel)', 'OS Ver=$($Script:OSVER)'" -Type 1 -LogFile $LogFile
        
        $Script:ExtractPackageShareDirCurrent = $null

	    # let's query for HP Driver Pack softpaqs with the HP CMSL
	    CMTraceLog -Message "Obtaining DriverPack Softpaq info" -Type $Script:TypeDebug -LogFile $LogFile
        $SoftPaq = Get-SoftpaqList -platform $lEntryProdCode -os $OS -osver $Script:OSVER
	    $HPDriverPack = $SoftPaq | Where-Object { $_.category -eq 'Manageability - Driver Pack' }
	    $HPDriverPack = $HPDriverPack | Where-Object { $_.Name -notmatch "Windows PE" }
	    $HPDriverPack = $HPDriverPack | Where-Object { $_.Name -notmatch "WinPE" }
       
	    if ($HPDriverPack) {

            # Get Current Driver CMPackage Version from CM for this Model - in case it exists already

		    CMTraceLog -Message "Looking for CM Driver Package" -Type $Script:TypeDebug -LogFile $LogFile
            Set-Location -Path "$($SiteCode):"
            $CMDriverPackage = Get-CMPackage -Name $lEntryModel -Fast
            Set-Location -Path "C:"

            # ------------------------------------------------------------------------
            # Check to see if:
            #    - there is a newer version at HP
            #    - (if versions match) the PackageID in CM matches the PackageID in the script
            if ($CMDriverPackage) {

                # see if the OS Versions of the CM Package match the version we are seeking now
                $lOSVerMatch = [string]$CMDriverPackage.PkgSourcePath -match $($Script:OSVER)

                if ( $lOSVerMatch ) {
                    if ($CMDriverPackage.version -ne $HPDriverPack.Version) {
                        CMTraceLog -Message "... CM Driver Pack version $($CMDriverPackage.version) needs updated to HP Version: $($HPDriverPack.Version)" -Type 2 -LogFile $LogFile
                    } 
                } else {
                    # The Version that the CM Package does NOT match the OS Version we are aspiring to
                    CMTraceLog -Message "... CM Package OS version does not match $($Script:OSVER) - $($CMDriverPackage.PkgSourcePath)" -Type 2 -LogFile $LogFile
                }
		    } else {
			        CMTraceLog -Message "... CM Driver Package NOT Found, HP version: $($HPDriverPack.Version)" -Type $ErrorType -LogFile $LogFile
            } # else if ( $CMDriverPackage )

            ################################################################################################
		    # Find out if the Driver Package step is in the TS
            # $TSDriverModule is the Tak Sequence containing all Dowlnoad DriverPack steps

            CMTraceLog -Message "Checking Package Step in Task Sequence" -Type $Script:TypeDebug -LogFile $LogFile
            $CMPackageStepFound = Update_TSPackageStep $TSDriverModule $CMDriverPackage $lEntryProdCode $true    # arg: $true = report only, don't update
            
		    ################################################################################################

        } else {
            $CMDriverPackage = $null
            $CMPackageStepFound = $false
		    #$EmailArray += "<font color=#8B0000>No Driver Pack Available for<b> $($HPModel.Model) </b>Product Code $($HPModel.ProdCode) via Internet </font><br>"
		    CMTraceLog -Message "... No HP Driver Pack Available for '$($lEntryModel)' Product Code '$($lEntryProdCode)'" -Type $ErrorType -LogFile $LogFile

	    } # else { # if ($HPDriverPack)

        ListView_FillInItem $pModelsList $_ $HPDriverPack $CMPackageStepFound

    } # $lCheckedItemsList | ForEach-Object  

    CMTraceLog -Message "Done Reporting" -Type 1 -LogFile $LogFile

} # Function CM_Report

#=====================================================================================
<#
    Function CM_Download
        This function will find out if a nee DriverPack is available for every product in the checkmarked list
        if available, it will download, unpack and populate the Software Package source
            then, will create or update the CM Software Package associated for the Hp Mode
            it will also create or update the Package entry (step) in the Driver Package Module Task Sequence 
    expects parameter 
        - Gui Model list table
        - list of checked entries from the table checkmarks
        - boolean downloadAndUpdate - for future use
#>
Function CM_Download {
    [CmdletBinding()]
	param( $pModelsList, $pCheckedItemsList, [bool]$pDownloadAndUpdate )
    
    CMTraceLog -Message "Downloading and Updating CM" -Type 1 -LogFile $LogFile

    CMTraceLog -Message "Checking for OS version: $($OSVER)" -Type $TypeDebug -LogFile $LogFile 

    # run thru every checked row # in the Models listView box
    $pCheckedItemsList | ForEach-Object  {

        # SubItems[0] is the checkmark column
        $lEntryProdCode = $pModelsList[1,$_].Value
        $lEntryModel = $pModelsList[2,$_].Value

        CMTraceLog -Message "Checking: table row:$($_); Model=$($lEntryModel); ProdCod=$($lEntryProdCode)" -Type 1 -LogFile $LogFile 
        
        $Script:ExtractPackageShareDirCurrent = $null                                 # This is set in Update_SoftpaqFolders
        # let's query for HP Driver Pack softpaqs with the HP CMSL
        $SoftPaq = Get-SoftpaqList -platform $lEntryProdCode -os $OS -osver $Script:OSVER
	    $HPDriverPack = $SoftPaq | Where-Object { $_.category -eq 'Manageability - Driver Pack' }
	    $HPDriverPack = $HPDriverPack | Where-Object { $_.Name -notmatch "Windows PE" }
	    $HPDriverPack = $HPDriverPack | Where-Object { $_.Name -notmatch "WinPE" }
	    
        if ( $HPDriverPack ) {

			CMTraceLog -Message "... Updating Softpaq folder" -Type 1 -LogFile $LogFile
            Update_SoftpaqFolders $HPDriverPack $lEntryModel  

            CMTraceLog -Message "... Looking for CM Driver Package for $($lEntryModel)" -Type $TypeDebug -LogFile $LogFile
            Set-Location -Path "$($SiteCode):"
            $CMDriverPackage = Get-CMPackage -Name $lEntryModel -Fast
            Set-Location -Path "C:"

            if ( $CMDriverPackage -eq $null ) {
                CMTraceLog -Message "... CM Package missing... Creating New" -Type 1 -LogFile $LogFile
                $CMDriverPackage = New-CMPackage -Name $lEntryModel -Manufacturer "HP"                

            } else {
                # see if the OS Versions of the CM Package match the version we are seeking now
                $lOSVerMatch = [string]$CMDriverPackage.PkgSourcePath -match $($Script:OSVER)

                if ( !$lOSVerMatch ) {
                    CMTraceLog -Message "... Package OS version does not match current selection" -Type $TypeDebug -LogFile $LogFile
                }
                CMTraceLog -Message "... found lEntryModel=$($lEntryModel) ; w/lCMDriverPackage.PackageID=$($CMDriverPackage.PackageID)" -Type $TypeDebug -LogFile $LogFile
                if ($CMDriverPackage.Version -eq $HPDriverPack.Version) {

			        CMTraceLog -Message "... CM Driver Package is up to date: version ""$($HPDriverPack.Version)""" -Type 1 -LogFile $LogFile 

                } else {
                     CMTraceLog -Message "... Driver Package version mismatch, updating to HP version ""$($HPDriverPack.Version)"", Softpaq ""$($HPDriverPack.Id)""" -Type 1 -LogFile $LogFile
                } # else if ($lCMDriverPackage.Version -eq $HPDriverPack.Version)

                ################################################################################################
			    # Now, let's work on the Driver Package STEP in the Driver Module TS

                CMTraceLog -Message "... Updating CM Task Sequence step with DriverPack Package - lCMDriverPackage.name=$($CMDriverPackage.name)" -Type 1 -LogFile $LogFile    
			    $CMPackageStepFound = Update_TSPackageStep $TSDriverModule $CMDriverPackage $lEntryProdCode $false
			    CMTraceLog -Message "... Package: ""$($CMDriverPackage.Name)""->$($CMDriverPackage.PackageID)" -Type $TypeDebug -LogFile $LogFile  

            } # else if ( $HPDriverPack )

            $CMDriverPackage = Set_HPDriverPackage $CMDriverPackage $HPDriverPack.ID $HPDriverPack.Version $ExtractPackageShareDirCurrent # set: name, Version, Path

        } else {
            $CMDriverPackage = $null
            $CMPackageStepFound = $false
        } # else if ( $HPDriverPack )

        ListView_FillInItem $pModelsList $_ $HPDriverPack $CMPackageStepFound

    } # $lCheckedItemsList | ForEach-Object 

    CMTraceLog -Message "Done!" -Type 1 -LogFile $LogFile

} # Function CM_Download

#=====================================================================================
<#
    Function CreateForm
    This is the MAIN function with a Gui that sets things up for the user
#>

Function CreateForm {
    
    Add-Type -assembly System.Windows.Forms

    $LeftOffset = 20
    $TopOffset = 20
    $FieldHeight = 20
    $FormWidth = 800
    $FormHeight = 600

    $CM_form = New-Object System.Windows.Forms.Form
    $CM_form.Text = "CM_Driverpack_Downloader v$($ScriptVersion)"
    $CM_form.Width = $FormWidth
    $CM_form.height = $FormHeight
    $CM_form.Autosize = $true
    $CM_form.StartPosition = 'CenterScreen'

    #----------------------------------------------------------------------------------
    #define a tooltip object
    $tooltips = New-Object System.Windows.Forms.ToolTip
    $ShowHelp={
        #display popup help
        #each value is the name of a control on the form.
    
        Switch ($this.name) {
            "OS_Selection"  {$tip = "Allowed Windows 10 Versions in Script"}
            "CM_Install_Path" {$tip = "System Path for SCCM Installation"}
            #"check2" {$tip = "Query Win32_Computersystem"}
            #"check3" {$tip = "Query Win32_BIOS"}
        } # Switch ($this.name)
        $tooltips.SetToolTip($this,$tip)
    } #end ShowHelp

    #----------------------------------------------------------------------------------
    $ActionComboBox = New-Object System.Windows.Forms.ComboBox
    $ActionComboBox.Width = 150
    $ActionComboBox.Location  = New-Object System.Drawing.Point($LeftOffset, $TopOffset)
    $ActionComboBox.DropDownStyle = "DropDownList"
    Foreach ($MenuItem in 'Report', 'Download', 'Setup CM Task Sequences') {
        [void]$ActionComboBox.Items.Add($MenuItem);
    }  
    $ActionComboBox.SelectedIndex = $defaultIndex

    $buttonGo = New-Object System.Windows.Forms.Button
    $buttonGo.Width = 30
    $buttonGo.Text = 'Go'
    $buttonGo.Location = New-Object System.Drawing.Point(($LeftOffset+155),($TopOffset-1))

    $buttonGo.add_click( {
        $Script:OSVER=$OSVERComboBox.Text
        if ( $Script:OSVER -in $Script:OSVALID ) {
            # get a list of all checked entries (row numbers, starting with 0)
            $lCheckedListArray = @()
            for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
                if ($dataGridView[0,$i].Value) {
                    $lCheckedListArray += $i
                } # if  
            } # for

            if ($Script:CMConnected -eq $true) {
                if ($ActionComboBox.SelectedItem -eq 'Report') {
                    CM_Report $dataGridView $lCheckedListArray
                } elseif ($ActionComboBox.SelectedItem -eq 'Download') {
                    CM_Download $dataGridView $lCheckedListArray $true
                } elseif ($ActionComboBox.SelectedItem -eq 'Setup CM Task Sequences') {
                    CM_CheckTS 
                } # elseif
            } else {
                CMTraceLog -Message "NO CM Connection!! Click on [Connect] button to start" -Type 1 -LogFile $LogFile
            } # else if ($Script:CMConnected -eq $true) 
        } else {
            CMTraceLog -Message "Windows 10 OS Version NOT Supported in Script: $($Script:OSVER) - Must be one of: $($Script:OSVALID)" -Type 3 -LogFile $LogFile  
        } # else
    } ) # $buttonGo.add_click
    
    $CM_form.Controls.AddRange(@($buttonGo, $ActionComboBox))

    #----------------------------------------------------------------------------------
    # Create OS and OS Version display fields - info from .ini file

    $OSTextLabel = New-Object System.Windows.Forms.Label
    $OSTextLabel.Text = "Windows 10:"
    $OSTextLabel.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+40))    # (from left, from top)
    $OSTextLabel.Size = New-Object System.Drawing.Size(70,25)                               # (width, height)
    $OSTextField = New-Object System.Windows.Forms.TextBox
    $OSVERComboBox = New-Object System.Windows.Forms.ComboBox
    $OSVERComboBox.Size = New-Object System.Drawing.Size(60,$FieldHeight)                  # (width, height)
    $OSVERComboBox.Location  = New-Object System.Drawing.Point(($LeftOffset+70), ($TopOffset+38))
    $OSVERComboBox.DropDownStyle = "DropDownList"
    $OSVERComboBox.Name = "OS_Selection"
    $OSVERComboBox.add_MouseHover($ShowHelp)
    
    Foreach ($MenuItem in $OSVALID) {
        [void]$OSVERComboBox.Items.Add($MenuItem);
    }  
    $OSVERComboBox.SelectedItem = $OSVER 

    $CM_form.Controls.AddRange(@($OSTextLabel,$OSVERComboBox))

    #----------------------------------------------------------------------------------
    # Create CM Connection Path Field (editable by default) and Test Connection Button

    $CMTextField = New-Object System.Windows.Forms.TextBox
    $CMTextField.top = 20
    $CMTextField.left = 20
    $CMTextField.Text = $CMInstall                                                          # populate with CM install Path
    $CMTextField.Multiline = $false 
    $CMTextField.Size = New-Object System.Drawing.Size(360,$FieldHeight)                    # (width, height)
    $CMTextField.ReadOnly = $true
    $CMTextField.Name = "CM_Install_Path"
    $CMTextField.add_MouseHover($ShowHelp)
    
    $DLPathLabel = New-Object System.Windows.Forms.Label
    $DLPathLabel.Text = "DL Path:"
    $DLPathLabel.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+20))    # (from left, from top)
    $DLPathLabel.Size = New-Object System.Drawing.Size(50,20)                               # (width, height)
    $DLPathLabel.TextAlign = "MiddleRight"
    $DLPathTextField = New-Object System.Windows.Forms.TextBox
    $DLPathTextField.Text = $DownloadSoftpaqDir
    $DLPathTextField.Multiline = $false 
    $DLPathTextField.location = New-Object System.Drawing.Point(($LeftOffset+50),($TopOffset+20))   # (from left, from top)
    $DLPathTextField.Size = New-Object System.Drawing.Size(310,$FieldHeight)                # (width, height)
    $DLPathTextField.ReadOnly = $true
    
    $SharePathLabel = New-Object System.Windows.Forms.Label
    $SharePathLabel.Text = "Share:"
    $SharePathLabel.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+40)) # (from left, from top)
    $SharePathLabel.Size = New-Object System.Drawing.Size(50,20)                            # (width, height)
    $SharePathLabel.TextAlign = "MiddleRight"
    $SharePathTextField = New-Object System.Windows.Forms.TextBox
    $SharePathTextField.Text = $ExtractPackageShareDir+"{OSVER}"
    $SharePathTextField.Multiline = $false 
    $SharePathTextField.location = New-Object System.Drawing.Point(($LeftOffset+50),($TopOffset+40)) # (from left, from top)
    $SharePathTextField.Size = New-Object System.Drawing.Size(310,$FieldHeight)             # (width, height)
    $SharePathTextField.ReadOnly = $true
    
    $CMGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMGroupBox.location = New-Object System.Drawing.Point(($LeftOffset+240),($TopOffset))     # (from left, from top)
    $CMGroupBox.Size = New-Object System.Drawing.Size(390,90)                              # (width, height)
    $CMGroupBox.text = "CM Path ($($IniFile)):"

    $CMGroupBox.Controls.AddRange(@($CMTextLabel, $CMTextField, $DLPathLabel, $DLPathTextField, $SharePathLabel, $SharePathTextField))

    $CMTestbutton = New-Object System.Windows.Forms.Button                                  # to be placed inside the GroupBox
    $CMTestbutton.Text = 'Connect'
    $CMTestbutton.top = $TopOffset+10
    $CMTestbutton.left = $LeftOffset+640
    $CMTestbutton.Width = 80
    
    $CMTestbutton.add_click( {
            $CMInstall = $CMTextField.Text
            $Script:CMConnected = Test_CMConnection
            CMTraceLog -Message "CM Connection established: $($CMConnected)" -Type 1 -LogFile $LogFile
        }
    ) # $buttonDone.add_click

    $DebugCheckBox = New-Object System.Windows.Forms.CheckBox
    $DebugCheckBox.Text = 'Debug Output'
    $DebugCheckBox.UseVisualStyleBackColor = $True
    $DebugCheckBox.location = New-Object System.Drawing.Point(($LeftOffset+640),($TopOffset+250))   # (from left, from top)
    $DebugCheckBox.add_click( {
            if ( $DebugCheckBox.checked ) {
                $Script:Debug_Output = $true
            } else {
                $Script:Debug_Output = $false
            }
        }
    ) # $DebugCheckBox.add_click

    $CM_form.Controls.AddRange(@($CMGroupBox, $CMTestbutton, $DebugCheckBox))
  
    #----------------------------------------------------------------------------------
    # Create Models (label and) list Checked Text box, and a clear list button to check/uncheck all entries
    # The ListView control allows columns to be used as fields in a row
    
    $ListViewWidth = 610
    $ListViewHeight = 160
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.location = New-Object System.Drawing.Point(($LeftOffset-10),($TopOffset))
    $dataGridView.height = $ListViewHeight
    $dataGridView.width = $ListViewWidth
    $dataGridView.ColumnHeadersVisible = $true                   # the column names becomes row 0 in the datagrid view
    $dataGridView.RowHeadersVisible = $false
    $dataGridView.SelectionMode = 'CellSelect'
    $dataGridView.AllowUserToAddRows = $False                    # Prevents the display of empty last row

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $CheckBoxColumn.width = 28

    [void]$DataGridView.Columns.Add($CheckBoxColumn) 
    
    $dataGridView.ColumnCount = 8                                # 1st column is checkbox column

    $dataGridView.Columns[1].Name = 'SysId'
    $dataGridView.Columns[1].Width = 40
    $dataGridView.Columns[1].DefaultCellStyle.Alignment = "MiddleCenter"
    $dataGridView.Columns[2].Name = 'Model'
    $dataGridView.Columns[2].Width = 210
    $dataGridView.Columns[3].Name = 'Available'
    $dataGridView.Columns[3].Width = 90
    $dataGridView.Columns[3].DefaultCellStyle.Alignment = "MiddleCenter"
    $dataGridView.Columns[4].Name = 'CM Ver'
    $dataGridView.Columns[4].Width = 50
    $dataGridView.Columns[4].DefaultCellStyle.Alignment = "MiddleCenter"
    $dataGridView.Columns[5].Name = 'CM OS'
    $dataGridView.Columns[5].Width = 50
    $dataGridView.Columns[5].DefaultCellStyle.Alignment = "MiddleCenter"
    $dataGridView.Columns[6].Name = 'CM Package'
    $dataGridView.Columns[6].Width = 80
    $dataGridView.Columns[6].DefaultCellStyle.Alignment = "MiddleCenter"
    $dataGridView.Columns[7].Name = 'Step'
    $dataGridView.Columns[7].Width = 40
    $dataGridView.Columns[7].DefaultCellStyle.Alignment = "MiddleCenter"

    # fill the listview box with all the HP Models listed in the ini script
    $HPModelsTable | ForEach-Object {
                        # populate 1st 3 columns: checkmark, SysID, Model Name
                        $row = @( $true, $_.ProdCode, $_.Model)         
                        [void]$dataGridView.Rows.Add($row)
                } # ForEach-Object

    # next 2 lines clear any selection from the initial data view
    $dataGridView.CurrentCell = $dataGridView[1,1]
    $dataGridView.ClearSelection()

    # Add a CheckBox on header (1st col)

    $CheckAll=New-Object System.Windows.Forms.CheckBox
    $CheckAll.AutoSize=$true
    $CheckAll.Left=9
    $CheckAll.Top=6
    $CheckAll.Checked = $true
    $CheckAll_Click={
        $state = $CheckAll.Checked
        for ($i = 0; $i -lt $dataGridView.RowCount; $i++) {
            $dataGridView[0,$i].Value = $state
        }
    } # $CheckAll_Click={

    $CheckAll.add_Click($CheckAll_Click)

    $dataGridView.Controls.AddRange(@($CheckAll))

    $CMModlesGroupBox = New-Object System.Windows.Forms.GroupBox
    $CMModlesGroupBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+90))     # (from left, from top)
    $CMModlesGroupBox.Size = New-Object System.Drawing.Size(($ListViewWidth+20),($ListViewHeight+30))       # (width, height)
    $CMModlesGroupBox.text = "HP Models"

    $CMModlesGroupBox.Controls.AddRange(@($dataGridView))

    $CM_form.Controls.AddRange(@($CMModlesGroupBox))
    
    #----------------------------------------------------------------------------------
    # Create Output Text Box at the bottom of the dialog

    $Script:TextBox = New-Object System.Windows.Forms.RichTextBox
    $TextBox.Name = $Script:FormOutTextBox                                          # named so other functions can output to it
    $TextBox.Multiline = $true
    $TextBox.Autosize = $false
    $TextBox.ScrollBars = "Both"
    $TextBox.WordWrap = $false
    $TextBox.location = New-Object System.Drawing.Point($LeftOffset,($TopOffset+280))            # (from left, from top)
    $TextBox.Size = New-Object System.Drawing.Size(($FormWidth-60),280)             # (width, height)

    $CM_form.Controls.AddRange(@($TextBox))

    #----------------------------------------------------------------------------------
    # Create Done/Exit Button at the bottom of the dialog

    $buttonDone = New-Object System.Windows.Forms.Button
    $buttonDone.Text = 'Done'
    $buttonDone.Location = New-Object System.Drawing.Point(($FormWidth-120),($FormHeight-50))    # (from left, from top)

    $buttonDone.add_click( {
            $CM_form.Close()
        }
    ) # $buttonDone.add_click
     
    $CM_form.Controls.AddRange(@($buttonDone))
    
    #----------------------------------------------------------------------------------
    # Finally, show the dialog on screen
    
    $Script:CMConnected = Test_CMConnection
    if ( !$Script:CMConnected ) {
        CMTraceLog -Message "Please Verify to CM Environment" -Type 1 -LogFile $LogFile
    }

    $CM_form.ShowDialog() | Out-Null

} # Function CreateForm


# --------------------------------------------------------------------------
# Start of Script
# --------------------------------------------------------------------------

#CMTraceLog -Message "Starting Script: $scriptName, version $ScriptVersion" -Type 1 -LogFile $LogFile

# Read the contents of the ini file, in case we need the data later
# WORK: TBD (idea to edit contents)
#$script:IniDataset = Get-Content -Path $IniFilePath -Raw

# Create the GUI and take over all actions, like Report and Download

CreateForm
<#
#>
