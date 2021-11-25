###
# Author:          Paul Jezek
# ScriptVersion:   v1.0.7, Nov 25, 2021
# Description:     WingetBridge Factory - Shared functions
# Compatibility:   MSI, NULLSOFT, INNO, BURN
# Please visit:    https://github.com/endpointmanager/wingetbridge-factory
#######

#######
# There is no configuration required within this script.
#######

function Get-PathFriendlyName
{
param(
    [string]$Name
)
    $tmp = $Name.Split(",")
    if ($tmp.Count -ige 2)
    {
        $Name = $tmp[0]
    }
    $Name = $Name.Trim(".")
    $Name = $Name.Replace("\", "")
    $Name = $Name.Replace("/", "")
    $Name = $Name.Replace(":", "")
    $Name = $Name.Replace("*", "")
    $Name = $Name.Replace("?", "")
    $Name = $Name.Replace("<", "")
    $Name = $Name.Replace(">", "")
    $Name = $Name.Replace("|", "")
    return $Name
}

function Get-WingetBridgeKeywordsAsString
{
param(
    [string[]]$Tags,
    [int]$maxlength
)
    $culture = Get-Culture
    $Seperator = $culture.TextInfo.ListSeparator
    $keywords = ""
    foreach ($tag in $manifest.Tags)
    {
        if ($tag.Length -lt 64)
        {
            $nextkeywords = "$($tag)$($Seperator)$($keywords)"
            if ($nextkeywords.Length -lt $maxlength)
            {
                $keywords = $nextkeywords
            }
        }
    }
    return $keywords
}

function Get-BestIconResolution
{
param (
    [string] $SourceFile
)
    $Resolutions = Save-WingetBridgeAppIcon -SourceFile $SourceFile -ValidateOnly
    $DefaultResolution = 32 #If no recommended entry found
    $RecommendedResolution = @(128, 96, 64, 48) #prefered 32bit Resolutions
    $AllowedResolutions = @(512, 256, 128, 64, 48, 32) #MEMCM does not support pngs hihger than 512x512
    if ($Resolutions.Count -gt 1)
    {
        foreach ($Resolution in $Resolutions)
        {            
            if ($Resolution.Bits -eq 32)
            {
                if ($Resolution.Width -in $RecommendedResolution)
                {                    
                    return $Resolution.Width;
                    break;
                }
            }
        }
        foreach ($Resolution in $Resolutions)
        {
            if ($Resolution.Width -in $AllowedResolutions)
            {                
                return $Resolution.Width;
                break;
            }
        }
    }
    else
    {
        if ($Resolutions.Count -eq 1) { return $Resolutions[0].Width; } else { return $DefaultResolution; }
    }
}

function Save-AsPng
{
param (
    [string] $SourceFile,
    [string] $TargetFile,
    [int] $Resolution
)
    if ($Global:VerboseMessages) { Write-Host "Save-AsPng ($SourceFile) to ($TargetFile) with a resolution of $($Resolution)x$($Resolution)" -ForegroundColor DarkCyan }
    if ($SourceFile)
    {
        $iconBmp = [System.Drawing.Bitmap]::FromFile($SourceFile)
    }
    
    $newbmp = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $iconBmp, $Resolution, $Resolution    
    
    $newbmp.Save($TargetFile, "png")
    $newbmp.Dispose()
    try
    {
        Test-Path $TargetFile | Out-Null        
    }
    catch
    {
        Write-Host "Cannot find new icon file.Error message: $($_.Exception.Message)" -ForegroundColor red
    }
}

function Get-PartOfInstructionLine #Helper-function for Get-InstallerDetailsFor...
{
param (
    [string] $Instruction
)
    $dde = $Instruction.Split(" ")
    $new = @()
    $quotefound = $false
    $tmp = ""
    foreach ($d in $dde)
    {
        if ($d.StartsWith("`""))
        {
            $quotefound = $true
        }
        else
        {
            if ($quotefound -eq $false) { $new += $d }
        }
        if ($quotefound)
        {
            if ($d.EndsWith("`""))
            {
                $tmp = "$tmp$d"
                $new += $tmp
                $tmp = ""
                $quotefound = $false
            }
            else
            {
                $tmp = "$tmp$d "
            }
        }
    }
    return $new
}

function Get-WindowsInstallerTableData {
	[CmdletBinding(SupportsShouldProcess=$True,DefaultParameterSetName="None")]
	PARAM (
	    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="MSI Database Filename",ValueFromPipeline=$true)]
	    [Alias("Database","Msi")]
		[ValidateScript({Test-Path "FileSystem::$_" -PathType 'Leaf'})]
		[System.IO.FileInfo]
	    $MsiDbPath
		,
	    [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="MST Database Filename",ValueFromPipeline=$true)]
	    [Alias("Transform","Mst")]
		[System.IO.FileInfo[]]
	    $MstDbPath
		,
	    [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="SQL Query",ValueFromPipeline=$true)]
	    [Alias("Query")]
		[String]
	    $Table
	)

	begin {
		# Add Required Type Libraries        
        Add-Type -Path "$Global:WindowsInstaller_dll"
	}
		
	process {
		# Open an MSI Database
		$Database = New-Object Microsoft.Deployment.WindowsInstaller.Database $MsiDbPath;
		# ApplyTransforms
		foreach ($MstDbPathEx In $MstDbPath) {
			if (Test-Path -Path "FileSystem::$MstDbPathEx" -PathType Leaf) {
				$Database.ApplyTransform($MstDbPathEx);
			}
		}
			
		# Create a View object		
		# OpenView returns data with type of [Microsoft.Deployment.WindowsInstaller.View]
		$_ColumnsView = $Database.OpenView("SELECT * FROM ``_Columns`` WHERE ``Table`` = '$($Table)'");
		if ($_ColumnsView) {
			# Execute the View object
			$_ColumnsView.Execute()
			# Place the objects in a PSObject
			$_Columns = @()
			$_ColumnsRow = $_ColumnsView.Fetch()
			while($_ColumnsRow -ne $null) {
				$hash = @{
					'Table' = $_ColumnsRow.GetString(1)
					'Number' = $_ColumnsRow.GetString(2)
					'Name' = $_ColumnsRow.GetString(3)
					'Type' = $_ColumnsRow.GetString(4)
				}
				$_Columns += New-Object -TypeName PSObject -Property $hash
					
				$_ColumnsRow = $_ColumnsView.Fetch()
			}

			$FieldNames = $_Columns | Select -ExpandProperty Name
		}
			
		if ($FieldNames) {
			# [Microsoft.Deployment.WindowsInstaller.View]
			$TableView = $Database.OpenView("SELECT * FROM ``$($Table)``");
			# Execute the View object
			$TableView.Execute()
			# Place the objects in a PSObject
			$Rows = @()
			# Fetch the first record
			$Row = $TableView.Fetch()
			while($Row -ne $null) {
				$hash = @{}
				foreach ($FieldName In $FieldNames) {
					$hash += @{
						$FieldName = $Row.Item($FieldName)
					}
				}
				$Rows += New-Object -TypeName PSObject -Property $hash
					
				# Fetch the next record
				$Row = $TableView.Fetch()
			}
			$Rows
		}
	}
		
	end {
		#Close the Database & View
		if ($_ColumnsView) {$_ColumnsView.Close();}
		if ($TableView) {$TableView.Close();}
		if ($Database) {$Database.Dispose();}
	}
}

function Get-WingetBridgeIcon
{
param (
    [string[]]$ExeShortCuts,
    [Hashtable]$InstallerDetails,
    [string]$InstallerType,
    [string]$CustomExtractor
)
    $ExeWithIcon = ""
    $SuggestedIconExists = $false

    foreach ($exe in $ExeShortCuts) #Search for suggested Icon
    {
        if ($InstallerDetails.DisplayIcon.contains("\$exe"))
        {
            $SuggestedIconExists = $true
            break
        }
    }
    
    foreach ($exe in $ExeShortCuts) #Create Icon From Shortcuts
    {                
        $ExeWithIcon = $exe        
        if ($ExeWithIcon -ne "")
        {
            Write-Host "Try to extract icon from `"$ExeWithIcon`" with `"$CustomExtractor`" ($InstallerType)" -ForegroundColor Magenta
            if ($InstallerType -eq "nullsoft")
            {
                $process = Start-Process -FilePath "$CustomExtractor" -ArgumentList @("e", "`"$InstallerFile`"", "-ir!`"$ExeWithIcon`"", "-o`"$($Global:TempDirectory)`"","-y") -WindowStyle Hidden -Wait -PassThru
            }
            if ($InstallerType -eq "msi")
            {
                $process = Start-Process -FilePath "$CustomExtractor" -ArgumentList @("x", "`"$InstallerFile`"", "$($Global:TempDirectory)\","`"$ExeWithIcon`"") -WindowStyle Hidden -Wait -PassThru
            }
            if ($InstallerType -eq "inno")
            {            
                $process = Start-Process -FilePath "$CustomInnoUnpSource" -ArgumentList @("`"$InstallerFile`"", "`"$ExeWithIcon`"", "-x", "-d`"$($Global:TempDirectory)`"", "-y") -WindowStyle Hidden -Wait -PassThru                
            }
            if ($process.ExitCode -eq 0)
            {                
                $extractedfiles = Get-ChildItem -Path $($Global:TempDirectory) -Filter $ExeWithIcon -Recurse
                foreach ($file in $extractedfiles)
                {
                    Move-Item -Path $file.FullName -Destination $($Global:TempDirectory) -Force
                }
                $CurrentIcon = ("$($Global:TempDirectory)\$ExeWithIcon").Replace(".exe",".ico")                
                if (Test-Path -Path "FileSystem::$($Global:TempDirectory)\$ExeWithIcon")
                {
                    $InstallerDetails.FileDetectionVersion = (Get-Item "$($Global:TempDirectory)\$ExeWithIcon").VersionInfo.FileVersionRaw
                
                    if (!(Test-Path -Path "FileSystem::$CurrentIcon"))
                    {
                        try
                        {
                            $nouptput = Save-WingetBridgeAppIcon -SourceFile "$($Global:TempDirectory)\$ExeWithIcon" -TargetIconFile $CurrentIcon
                        }
                        catch
                        {
                            Write-Host "We were not able to extract an icon from the shortcut `"$ExeWithIcon`"" -ForegroundColor Yellow
                        } #could be an executable that does not contain an icon
                    }
                }
                else
                {
                    Write-Host "We were not able to extract the shortcut `"$ExeWithIcon`" from installer" -ForegroundColor Yellow
                }
                if (Test-Path -Path "FileSystem::$CurrentIcon")
                {
                    $InstallerDetails.TemporaryIconFile = $CurrentIcon
                    $InstallerDetails.BestIconResolution = Get-BestIconResolution -SourceFile "$($Global:TempDirectory)\$ExeWithIcon"
                    if (($InstallerDetails.DisplayIcon.contains("\$exe")) -or (!($SuggestedIconExists))) #this is the suggested icon (or there is no suggested icon)
                    {                        
                        break
                    }
                }
            } else
            {
                Write-Host "Failed to extract shortcut from installer (to get it's icon)" -ForegroundColor Red
            }
        } else { Write-Host "no icon found" }
    }
    return $InstallerDetails
}

function Test-InstallerAuthenticodeSignature
{
param (
    [string]$InstallerFile
)
    $isValidSignature = $false
    $AuthenticodeSignature = $null
    try
    {
        $AuthenticodeSignature = (Get-AuthenticodeSignature -FilePath "FileSystem::$InstallerFile")
    }
    catch {}
    if ($AuthenticodeSignature -ne $null) {
        if ($AuthenticodeSignature.Status -eq 0) { $isValidSignature = $true } #0 means "valid"
    }
    if (!($isValidSignature)) { Write-Host "WARNING: The signature of the file is not valid or missing! ($InstallerFile)" -ForegroundColor Red }
    return $isValidSignature
}

function Get-InstallerDetailsForMSI
{
param (
    [string]$InstallerFile,
    [string]$CustomMsiLessSource,
    [bool]$ExtractIcon
)
    
    $installerDetails = @{
        Name = ""
        Publisher = ""
        InstallDir = ""
        UninstallRegKey = ""
        DisplayIcon = ""
        DisplayName = ""
        DisplayVersion = ""        
        UninstallString = ""        
        TemporaryIconFile = "" #extracted in temp-directory
        FileDetection = ""
        FileDetectionVersion = ""
        BestIconResolution = 0
        ProductCode = ""
        UpgradeCode = ""
        FileCount = 0
        BinaryDataCount = 0
    }

    if (!(Test-Path -Path "FileSystem::$InstallerFile")) { Write-Host "Installer not found ($InstallerFile)" -ForegroundColor Red } else {
        $IsSigned = Test-InstallerAuthenticodeSignature -InstallerFile "$InstallerFile"
        $properties = Get-WindowsInstallerTableData -MsiDbPath $InstallerFile -Table "Property"

        if ($properties -ne $null)
        {      
            $installerDetails['ProductCode'] = ($properties.Where{ $_.Property -eq "ProductCode" }).Value
            $installerDetails['UpgradeCode'] = ($properties.Where{ $_.Property -eq "UpgradeCode" }).Value
            $installerDetails['Publisher'] = ($properties.Where{ $_.Property -eq "Manufacturer" }).Value
            $installerDetails['Name'] = ($properties.Where{ $_.Property -eq "ProductName" }).Value
            $installerDetails['DisplayVersion'] = ($properties.Where{ $_.Property -eq "ProductVersion" }).Value

            $shortcuts = Get-WindowsInstallerTableData -MsiDbPath $InstallerFile -Table "Shortcut"
            $ExeShortCuts = @()
            $shortcutsToIgnore = @("cmd.exe", "uninstall.exe", "iexplore.exe")
            foreach ($shortcut in $shortcuts)
            {                
                if ($shortcut.Target.EndsWith(".exe"))
                {
                    $ShortcutTarget = $shortcut.Target                    
                    $parts = $ShortcutTarget.Split("]") #Eliminate any preciding Variables like [APPLICATIONFOLDER], [INSTALLFOLDER] or [INSTALLDIR]
                    if ($parts.Count -gt 1)
                    {
                        $ShortcutTarget = $parts[1]
                    }
                    $ExeFile = [io.path]::GetFileName($ShortcutTarget)
                    if ($ExeFile.ToLower() -notin $shortcutsToIgnore)
                    {
                        if (!($ExeShortCuts.Contains($ExeFile))) { $ExeShortCuts += $ExeFile }
                        
                    }
                }
            }

            $filesInMsi = Get-WindowsInstallerTableData -MsiDbPath $InstallerFile -Table "File"
            $installerDetails['FileCount'] = $filesInMsi.Count
            if ($ExeShortCuts.Count -eq 0) #Try to get any executable from the MSI (could be a MSI-legacy-setup-wrapper)
            {
                foreach ($file in $filesInMsi)
                {                                  
                    $currentfilename = $file.Filename                    
                    $parts = $currentfilename.Split("|")
                    if ($parts.Count -gt 1)
                    {
                        $currentfilename = $parts[1]
                    }
                    if ($currentfilename.EndsWith(".exe"))
                    {
                        $ExeFile = $currentfilename
                        if ($ExeFile.ToLower() -notin $shortcutsToIgnore)
                        {
                            if (!($ExeShortCuts.Contains($ExeFile))) { $ExeShortCuts += $ExeFile }                        
                        }
                    }                
                }
            }

            $binariesInMsi = Get-WindowsInstallerTableData -MsiDbPath $InstallerFile -Table "Binary"
            $installerDetails['BinaryDataCount'] = $binariesInMsi.Count
            if (($ExtractIcon) -or ($installerDetails['UninstallRegKey'] -eq ""))
            {
                $InstallerDetails = Get-WingetBridgeIcon -InstallerDetails $InstallerDetails -InstallerType "msi" -CustomExtractor $CustomMsiLessSource -ExeShortCuts $ExeShortCuts
            }
            return $InstallerDetails
        }
        else
        {
            Write-Host "Setup does not look like an msi-installer!" -ForegroundColor Red
        }
    }
}

function Get-InstallerDetailsForBURN
{
param (
    [string]$InstallerFile,
    [string]$Custom7zSource,    
    [bool]$ExtractIcon    
)

    $installerDetails = @{
        Name = ""
        Publisher = ""
        InstallDir = ""
        UninstallRegKey = ""
        DisplayIcon = ""
        DisplayName = ""
        DisplayVersion = ""        
        UninstallString = ""        
        TemporaryIconFile = "" #extracted in temp-directory
        FileDetection = ""
        FileDetectionVersion = ""
        BestIconResolution = 0
        RegistrationId  = "" #BURN-specific
    }

    if (!(Test-Path -Path "FileSystem::$InstallerFile")) { Write-Host "Installer not found ($InstallerFile)" -ForegroundColor Red } else {
        $IsSigned = Test-InstallerAuthenticodeSignature -InstallerFile "$InstallerFile"

        #Verify if 7-Zip exists, to extract manifest from BURN/Wix-Bootstrapper
        if (($Custom7zSource -eq $null) -or ($Custom7zSource -eq "")) { $Custom7zSource = "C:\Program Files\7-Zip\7z.exe" }
        Write-Host "Verify existence of 7-Zip [$Custom7zSource]" -ForegroundColor Magenta #for BURN, no specific version of 7-zip is required (tested with 15.05 beta)
        if (!(Test-Path -Path "FileSystem::$Custom7zSource"))
        {        
            Write-Host "Please install 7-Zip (Version 15.05, 64bit is recommended). It is required to extract BURN manifest" -ForegroundColor Red
            break
        }

        #Create subdirectory inside the temp folder        
        $Foldername = [io.path]::GetFileName($InstallerFile).TrimEnd(".exe")
        $nf = New-Item -ItemType Directory -Path "FileSystem::$Global:TempDirectory\$Foldername" -Force                        
        Start-Process -FilePath "$Custom7zSource" -ArgumentList @("x", "`"$InstallerFile`"", "-i!`"0`"", "-o`"$($Global:TempDirectory)\$Foldername\`"", "-y") -WindowStyle Hidden -Wait
        
        if (Test-Path -Path "FileSystem::$($Global:TempDirectory)\$Foldername\0")
        {
            $BURN_manifest = New-Object xml            
            $BURN_manifest.Load( (Convert-Path "FileSystem::$($Global:TempDirectory)\$Foldername\0"))
            if ($BURN_manifest.BurnManifest -ne $null)
            {
                $installerDetails['RegistrationId']  = $BURN_manifest.BurnManifest.Registration.Id
                $installerDetails['UninstallString']  = "C:\ProgramData\Package Cache\$($BURN_manifest.BurnManifest.Registration.Id)\$($BURN_manifest.BurnManifest.Registration.ExecutableName)"
                if ($BURN_manifest.BurnManifest.Registration.Arp -ne $null)
                {
                    $installerDetails['Name']  = $BURN_manifest.BurnManifest.Registration.Arp.DisplayName
                    $installerDetails['DisplayName']  = $BURN_manifest.BurnManifest.Registration.Arp.DisplayName
                    $installerDetails['DisplayVersion']  = $BURN_manifest.BurnManifest.Registration.Arp.DisplayVersion
                    $installerDetails['Publisher']  = $BURN_manifest.BurnManifest.Registration.Arp.Publisher
                }
                $installerDetails['UninstallRegKey']  = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($BURN_manifest.BurnManifest.Registration.Id)" #seems to be always registered in Wow6432Node?!
            }
            else
            {
                Write-Host "Setup does not look like a BURN-installer!" -ForegroundColor Red
            }
        }
        else
        {
            Write-Host "Setup does not look like a BURN-installer!" -ForegroundColor Red
        }

        #At the moment we extract the icon from the installer.exe
        if ($ExtractIcon)
        {
            #$installerDetails = Get-WingetBridgeIcon -InstallerDetails $InstallerDetails -InstallerType "nullsoft" -CustomExtractor $Custom7zSource -ExeShortCuts @( "$InstallerFile" )
            $CurrentIcon = "$($Global:TempDirectory)\$Foldername\$([io.path]::GetFileName($InstallerFile).Replace(".exe",".ico"))"
            if (!(Test-Path -Path "FileSystem::$CurrentIcon"))
            {
                try
                {
                    $nouptput = Save-WingetBridgeAppIcon -SourceFile $InstallerFile -TargetIconFile $CurrentIcon
                }
                catch
                {
                    Write-Host "We were not able to extract an icon from `"$InstallerFile`"" -ForegroundColor Yellow
                } #could be an executable that does not contain an icon
            }
            if (Test-Path -Path "FileSystem::$CurrentIcon")
            {
                $InstallerDetails.TemporaryIconFile = $CurrentIcon
                $InstallerDetails.BestIconResolution = Get-BestIconResolution -SourceFile "$InstallerFile"
            }
        }
        

        #Cleanup temp directory
        Write-Host "cleanup tempfolder" -ForegroundColor Magenta
        if (!(Test-Path -Path "FileSystem::$($nf.FullName)")) { Remove-Item -Recurse "FileSystem::$($nf.FullName)" }
        return $installerdetails
    }
}

function Get-InstallerDetailsForInno
{
param (
    [string]$InstallerFile,
    [string]$CustomInnoUnpSource,
    [bool]$ExtractIcon
)
    function Get-BestGuessPath
    {
        param (
        [string]$InnoValue,
        [bool]$is64Bit
        )
        $tmp = $InnoValue                
        $tmp = $tmp.Replace("{pf32}","%ProgramFiles(x86)%")
        $tmp = $tmp.Replace("{pf64}","%ProgramFiles%")
        $tmp = $tmp.Replace("{code:DefDirRoot}","{pf}")
        $tmp = $tmp.Replace("{autopf}","{pf}")
        $tmp = $tmp.Replace("{commonpf}","{pf}")
        $tmp = $tmp.Replace("{commonpf32}","{pf}")
        $tmp = $tmp.Replace("{commonpf64}","{pf}")
        if ($is64Bit)
        {
            $tmp = $tmp.Replace("{pf}","%ProgramFiles%")            
        }
        else
        {
            $tmp = $tmp.Replace("{pf}","%ProgramFiles(x86)%")            
        }
        return $tmp
    }
    $installerDetails = @{
        Name = ""
        Publisher = ""
        InstallDir = ""
        UninstallRegKey = ""
        DisplayIcon = ""
        DisplayName = ""
        DisplayVersion = ""
        UninstallString = ""
        TemporaryIconFile = "" #extracted in temp-directory
        FileDetection = ""
        FileDetectionVersion = ""
        BestIconResolution = 0
        Is64bit = $false
    }

    if (!(Test-Path -Path "FileSystem::$InstallerFile")) { Write-Host "Installer not found ($InstallerFile)" -ForegroundColor Red } else {
        $IsSigned = Test-InstallerAuthenticodeSignature -InstallerFile "$InstallerFile"

        #Verify if correct InnoUnp version is installed to extract ISS (INNO Setup Script)        
        if ($Global:VerboseMessages) { Write-Host "Verify version of InnoUnp [$CustomInnoUnpSource]" -ForegroundColor DarkCyan }
        $InnoUnpInfo = Get-Item "FileSystem::$CustomInnoUnpSource"
        if (!($InnoUnpInfo.VersionInfo.ProductVersion -ge "0.50"))
        {            
            Write-Host "Please install InnoUnp Version 0.50 or higher. It is required to get inno instructions" -ForegroundColor Red
            break
        }        

        $expectedISS = "$($Global:TempDirectory)\install_script.iss"
        if (Test-Path -Path "FileSystem::$expectedISS") { Remove-Item "FileSystem::$expectedISS" -Force }

        Start-Process -FilePath "$CustomInnoUnpSource" -ArgumentList @("`"$InstallerFile`"", "install_script.iss", "-x", "-d`"$($Global:TempDirectory)`"", "-y") -WindowStyle Hidden -Wait

        if (Test-Path -Path "FileSystem::$expectedISS")
        {
            $array = gc "$expectedISS"            
            $currentSection = ""
            $appid = ""
            $CreateAppDir = $true
            $DefaultDirName = ""
            
            #Get Variables
            for ( $i=0; $i -le ($array.length - 1); $i++) {
                if (($array[$i].StartsWith('[')) -and ($array[$i].EndsWith(']'))) {
                    $currentSection = $array[$i]                    
                }
                if ($currentSection -eq "[Setup]")
                {
                    if ($array[$i].StartsWith('ArchitecturesInstallIn64BitMode=')) {
                        $CurrentArr = $array[$i].Split("=")
                        if ($CurrentArr[1].Trim() -eq "x64")
                        {                            
                            $installerDetails['is64bit'] = $true
                        }
                    }                    
                    if ($array[$i].StartsWith('CreateAppDir=')) { #we need detect this, so we wont use unins000.exe if no AppDir exists
                        $CurrentArr = $array[$i].Split("=")
                        if ($CurrentArr[1].ToLower().Trim() -eq "no")
                        {
                            Write-Host "no appdir will be created!" -ForegroundColor Red
                            $CreateAppDir = $false
                        }
                    }
                    if ($array[$i].StartsWith('DefaultDirName=')) {
                        $CurrentArr = $array[$i].Split("=")
                        $DefaultDirName = $CurrentArr[1].Trim().Replace("`"","")
                    }
                    if ($array[$i].StartsWith('AppId=')) {
                        $CurrentArr = $array[$i].Split("=")
                        $appid = $CurrentArr[1].Trim().Replace("{{","{")
                        if (($appid.StartsWith("{code:GetAppId|")) -and ($appid.EndsWith("}")))
                        {
                            $appid = $appid.Replace("{code:GetAppId|","")
                            $appid = $appid.Trim("}")
                        }
                    }
                    if ($array[$i].StartsWith('AppName=')) {
                        $CurrentArr = $array[$i].Split("=")
                        $installerDetails['Name'] = $CurrentArr[1].Trim()
                    }
                    if ($array[$i].StartsWith('AppPublisher=')) {
                        $CurrentArr = $array[$i].Split("=")
                        $installerDetails['Publisher'] = $CurrentArr[1].Trim()
                    }
                    if ($array[$i].StartsWith('AppVerName=')) {
                        $CurrentArr = $array[$i].Split("=")
                        $installerDetails['DisplayName'] = $CurrentArr[1].Trim()
                    }
                    if ($array[$i].StartsWith('AppVersion=')) {
                        $CurrentArr = $array[$i].Split("=")
                        $installerDetails['DisplayVersion'] = $CurrentArr[1].Trim()
                    }
                    if ($array[$i].StartsWith('UninstallDisplayIcon=')) {
                        $CurrentArr = $array[$i].Split("=")
                        $installerDetails['DisplayIcon'] = $CurrentArr[1].Trim()
                    }
                }
            }

            #Calculate InstallDir for machine-targeted installers (using a "best guess method")
            $installerDetails['DisplayIcon'] = Get-BestGuessPath -InnoValue $installerDetails['DisplayIcon'] -is64Bit $installerDetails['is64bit']
            $DefaultDirName = Get-BestGuessPath -InnoValue $DefaultDirName -is64Bit $installerDetails['is64bit']
            if ((!$DefaultDirName.Contains("{")) -and (!$DefaultDirName.Contains("}"))) #we might have not fully resolved the varialbe
            {
                $installerDetails['InstallDir'] = $DefaultDirName
                $installerDetails['DisplayIcon'] = $($installerDetails['DisplayIcon']).Replace("{app}", $installerDetails['InstallDir'])
            }            


            #Calculate fields for uninstaller (very specific to inno)
            if ($appid -ne "")
            {
                $installerDetails['UninstallRegKey'] = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($appid)_is1"
                if (($installerDetails['InstallDir'] -ne "") -and ($CreateAppDir -eq $true)) #otherwise it should be uninstalled through the installer
                {
                    $installerDetails['UninstallString'] = "$($installerDetails['InstallDir'])\unins000.exe"
                }
            }            

            $UninstallExe = [io.path]::GetFileName($installerDetails['UninstallString'])
            $ExeShortCuts = @()
            $shortcutsToIgnore = @("cmd.exe", "uninstall.exe", "iexplore.exe")
            for ( $i=0; $i -le ($array.length - 1); $i++) {
                if (($array[$i].StartsWith('[')) -and ($array[$i].EndsWith(']'))) {
                    $currentSection = $array[$i]                    
                }
                if (($currentSection -eq "[Icons]") -and ($array[$i].Contains(".exe"))) {                    
                    $CurrentArr = $array[$i].Split(";")
                    foreach ($arr in $CurrentArr)
                    {
                        $CurrentParam = $arr.Trim()
                        if ($CurrentParam.StartsWith('Filename:'))
                        {
                            $ValueOfParam = Get-PartOfInstructionLine -Instruction $CurrentParam
                            if ($ValueOfParam.Count -eq 2)
                            {
                                $ExeFile = [io.path]::GetFileName($($ValueOfParam[1]).Trim("`""))
                                if (($ExeFile.ToLower().EndsWith(".exe")) -and ($ExeFile.ToLower() -notin $shortcutsToIgnore))
                                {
                                    if (!($ExeShortCuts.Contains($ExeFile))) { $ExeShortCuts += $ExeFile }
                                    $installerDetails['FileDetection'] = $($ValueOfParam[1]).Trim("`"") #Store full path for better file Detection
                                    if ($installerDetails['InstallDir'] -ne "")
                                    {
                                        $installerDetails['FileDetection'] = $installerDetails['FileDetection'].Replace("{app}", $installerDetails['InstallDir'])
                                    }
                                }
                            }
                        }
                    }
                }
            }

            #Search for Filename-Mappings (source- vs. target-filename) and alternative Executables
            $FilenameMapping = @{}
            for ( $i=0; $i -le ($array.length - 1); $i++) {
                if (($array[$i].StartsWith('[')) -and ($array[$i].EndsWith(']'))) {
                    $currentSection = $array[$i]
                }
                if (($currentSection -eq "[Files]") -and ($array[$i].Contains(".exe"))) {
                    $CurrentArr = $array[$i].Split(";")
                    $CurrentSourceParam = ""
                    $CurrentDestNameValue = ""
                    foreach ($arr in $CurrentArr)
                    {                        
                        $CurrentParam = $arr.Trim().Split(":")                        
                        if ($CurrentParam.Count -eq 2)
                        {                                                                                    
                            if ($CurrentParam[0] -eq "Source")
                            {
                                $CurrentSourceValue = [io.path]::GetFileName($($CurrentParam[1].Trim()).Replace("`"",""))
                                if ($ExeShortCuts.Count -eq 0) #Search for alternatives (*.exe matching AppName)
                                {
                                    $PossibleExeName = $($installerDetails['Name']).Replace(" ", "")
                                    if ($CurrentSourceValue.Contains("$PossibleExeName.exe"))
                                    {                                        
                                        $ExeShortCuts += $CurrentSourceValue
                                    }
                                }
                            }
                            if ($CurrentParam[0] -eq "DestName")
                            {
                                $CurrentDestNameValue = $($CurrentParam[1].Trim()).Replace("`"","")
                                $FilenameMapping[$CurrentDestNameValue] = $CurrentSourceValue
                            }
                        }                        
                    }                    
                }
            }

            if ($FilenameMapping.Count -gt 0)
            {
                $tempExeShortCuts = @()                
                foreach ($ExeFile in $ExeShortCuts)
                {
                    if ($FilenameMapping[$ExeFile] -ne $null) { $ExeFile = $FilenameMapping[$ExeFile] }
                    if (!($tempExeShortCuts.Contains($ExeFile))) { $tempExeShortCuts += $ExeFile }
                }
                $ExeShortCuts = $tempExeShortCuts
            }

            if (($ExtractIcon) -or ($installerDetails['UninstallRegKey'] -eq ""))
            {
                $installerDetails = Get-WingetBridgeIcon -InstallerDetails $InstallerDetails -InstallerType "inno" -CustomExtractor $CustomInnoUnpSource -ExeShortCuts $ExeShortCuts
            }
            return $installerDetails            
        }
        else
        {
            Write-Host "Setup does not look like an inno-installer!" -ForegroundColor Red
        }
    }
}

function Get-InstallerDetailsForNullsoft
{
param (
    [string]$InstallerFile,
    [string]$Custom7zSource,
    [bool]$ExtractIcon
)

    $installerDetails = @{
        Name = ""
        Publisher = ""
        InstallDir = ""
        UninstallRegKey = ""
        DisplayIcon = ""
        DisplayName = ""
        DisplayVersion = ""        
        UninstallString = ""        
        TemporaryIconFile = "" #extracted in temp-directory
        FileDetection = ""
        FileDetectionVersion = ""
        BestIconResolution = 0
    }

    if (!(Test-Path -Path "FileSystem::$InstallerFile")) { Write-Host "Installer not found ($InstallerFile)" -ForegroundColor Red } else {
        $IsSigned = Test-InstallerAuthenticodeSignature -InstallerFile "$InstallerFile"

        #Verify if correct 7z version is installed to extract NSIS (nullsoft)
        if (($Custom7zSource -eq $null) -or ($Custom7zSource -eq "")) { $Custom7zSource = "C:\Program Files\7-Zip\7z.exe" }
        if ($Global:VerboseMessages) { Write-Host "Verify version of 7-Zip [$Custom7zSource]" -ForegroundColor DarkCyan }
        $7zInfo = Get-Item "FileSystem::$Custom7zSource"
        if (!($7zInfo.VersionInfo.FileVersion -like "15.05 beta"))
        {
            Write-Host "Please install 7z Version 15.05 (64bit is recommended) It is required to get nullsoft instructions" -ForegroundColor Red            
            break
        }
                
        $expectedNSIS = "$($Global:TempDirectory)\``[NSIS``].nsi"
        if (Test-Path -Path "FileSystem::$expectedNSIS") { Remove-Item "FileSystem::$expectedNSIS" -Force }

        Start-Process -FilePath "$Custom7zSource" -ArgumentList @("x", "`"$InstallerFile`"", "-i!*.nsi", "-o`"$($Global:TempDirectory)`"") -WindowStyle Hidden -Wait

        if (Test-Path -Path "FileSystem::$expectedNSIS")
        {
            $array = gc "$expectedNSIS"

            #Get Variables
            for ( $i=0; $i -le ($array.length - 1); $i++) {
                if ($array[$i].StartsWith('Name')) {
                    $DisplayVersionArr = Get-PartOfInstructionLine -Instruction $array[$i].Trim()
                    $installerDetails['Name'] = $DisplayVersionArr[1].Trim("`"") 
                }
                if ($array[$i].StartsWith('InstallDir ')) {
                    $InstallDirArr = Get-PartOfInstructionLine -Instruction $array[$i].Trim()
                    $installerDetails['InstallDir'] = $InstallDirArr[1].Trim("`"")                     
                }
                if ($array[$i].ToLower().StartsWith('page custom')) {
                    break
                }
            }

            #Get Uninstaller-Variables
            for ( $i=0; $i -le ($array.length - 1); $i++) {
                if (($array[$i].Contains("Software\Microsoft\Windows\CurrentVersion\Uninstall")) -and
                (($array[$i].Contains("WriteRegStr HKLM")) -or ($array[$i].Contains("WriteRegStr SHCTX")) -or
                ($array[$i].Contains("WriteRegExpandStr HKLM")) -or ($array[$i].Contains("WriteRegExpandStr SHCTX")) )) { #SHCTX is HKLM if Shell-Context is Machine-based, otherwise it's HKCU which we are not targeting in this script
                    $DisplayVersionArr = Get-PartOfInstructionLine -Instruction $array[$i].Trim()
                    if ($DisplayVersionArr.Count -eq 5)
                    {
                        $ValueName = $DisplayVersionArr[3] #ValueName
                        if ($ValueName -eq "DisplayVersion")
                        {
                            $RegKey = $DisplayVersionArr[2] #RegKey
                            $installerDetails['UninstallRegKey'] = $RegKey.Trim("`"")

                            $DisplayVersion = $DisplayVersionArr[4] #Value
                            $installerDetails['DisplayVersion'] = $DisplayVersion
                        }
                        if ($ValueName -eq "DisplayIcon")
                        {
                            $DisplayIcon = $DisplayVersionArr[4] #Value
                            $installerDetails['DisplayIcon'] = $DisplayIcon 
                        }
                        if ($ValueName -eq "DisplayName")
                        {
                            $DisplayName = $DisplayVersionArr[4] #Value
                            $installerDetails['DisplayName'] = $DisplayName.Trim("`"")
                        }
                        if ($ValueName -eq "Publisher")
                        {
                            $Publisher = $DisplayVersionArr[4] #Value
                            $installerDetails['Publisher'] = $Publisher.Trim("`"")
                        }
                        if ($ValueName -eq "UninstallString")
                        {
                            $UninstallString = $DisplayVersionArr[4] #Value
                            if (($UninstallString.Contains(".exe")) -and ($UninstallString.Contains("`$INSTDIR"))) #Check it actually contains the uninstaller, and not a variable or register
                            {
                                $installerDetails['UninstallString'] = $UninstallString.Replace("$\`"", "")
                            }
                        }
                    }
                }
                if ($array[$i].Contains("StrCpy `$INSTDIR")) #might be useful for custom x86/x64 Handlers
                {
                    $StrCpyArr = Get-PartOfInstructionLine -Instruction $array[$i].Trim()                    
                    if ($StrCpyArr.Count -eq 3)
                    {
                        if ($StrCpyArr[2].Contains("`$PROGRAMFILES")) #only accept 1 of 3 ProgramFiles-Variables (neutral, 32, 64)
                        {
                            $installerDetails['InstallDir'] = $StrCpyArr[2]
                        }
                    }
                }
            }

            $UninstallExe = [io.path]::GetFileName($installerDetails['UninstallString'])            
            $ExeShortCuts = @()
            $shortcutsToIgnore = @("cmd.exe", "uninstall.exe", "iexplore.exe")
            for ( $i=0; $i -le ($array.length - 1); $i++) {
                if (($array[$i].Contains("CreateShortCut ")) -and ($array[$i].Contains(".exe"))) {
                    $DisplayVersionArr = Get-PartOfInstructionLine -Instruction $array[$i].Trim()                    
                    if ($DisplayVersionArr[2].EndsWith(".exe") -and ($DisplayVersionArr[2].ToLower() -ne $UninstallExe.ToLower()))
                    {                        
                        $ExeFile = [io.path]::GetFileName($DisplayVersionArr[2])
                        if ($ExeFile.ToLower() -notin $shortcutsToIgnore)
                        {
                            if (!($ExeShortCuts.Contains($ExeFile))) { $ExeShortCuts += $ExeFile }
                            $installerDetails['FileDetection'] = $DisplayVersionArr[2] #Store full path for better file Detection
                            $installerDetails['FileDetection'] = $installerDetails['FileDetection'].Replace("`$INSTDIR", $installerDetails['InstallDir'])
                            $installerDetails['FileDetection'] = $installerDetails['FileDetection'].Replace("`$PROGRAMFILES64", "%ProgramFiles%")
                            $installerDetails['FileDetection'] = $installerDetails['FileDetection'].Replace("`$PROGRAMFILES32", "%ProgramFiles(x86)%")
                            $installerDetails['FileDetection'] = $installerDetails['FileDetection'].Replace("`$PROGRAMFILES", "%ProgramFiles%")
                        }
                    }
                }
                if ($array[$i].Contains("WriteUninstaller")) #Fallback for UninstallString (if not set yet, from "Software\Microsoft\Windows\CurrentVersion\Uninstall")
                {
                    if ($installerDetails['UninstallString'] -eq "") #not set from "Software\Microsoft\Windows\CurrentVersion\Uninstall"
                    {
                        $WriteUninstArr = Get-PartOfInstructionLine -Instruction $array[$i].Trim()
                        if ($WriteUninstArr.Count -gt 1)
                        {
                            $UninstallString = $WriteUninstArr[1]
                            if (($UninstallString.Contains(".exe")) -and ($UninstallString.Contains("`$INSTDIR")))
                            {
                                $installerDetails['UninstallString'] = $UninstallString.Replace("$\`"", "")
                            }
                        }
                    }
                }
            }
            if (($ExtractIcon) -or ($installerDetails['UninstallRegKey'] -eq ""))
            {
                $installerDetails = Get-WingetBridgeIcon -InstallerDetails $InstallerDetails -InstallerType "nullsoft" -CustomExtractor $Custom7zSource -ExeShortCuts $ExeShortCuts
            }
            return $installerDetails
        }
        else
        {
            Write-Host "Setup does not look like a nullsoft-installer!" -ForegroundColor Red
        }
    }
}

function Remove-WingetBridgeFactoryTempFiles
{
param (
    [string]$TempDirectory,
    [bool]$FullCleanup
)
    try
    {
        Remove-Item -Path "$TempDirectory\*.exe" -Filter "*.exe" -Force
        Remove-Item -Path "$TempDirectory\*.nsi" -Filter "*.nsi" -Force
        Remove-Item -Path "$TempDirectory\*.iss" -Filter "*.iss" -Force
        if ($FullCleanup)
        {
            Remove-Item -Path "$TempDirectory\*.png" -Filter "*.png" -Force
            Remove-Item -Path "$TempDirectory\*.ico" -Filter "*.ico" -Force
            gci "$TempDirectory" -directory -recurse | Where { (gci $_.fullName).count -eq 0 } | select -expandproperty FullName | Foreach-Object { Remove-Item $_ } #Delete empty folders
        }
    } catch {}
}