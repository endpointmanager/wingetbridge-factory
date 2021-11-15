###
# Author:          Paul Jezek
# ScriptVersion:   v1.0.3, Nov 14, 2021
# Description:     WingetBridge Factory - Shared functions
# Compatibility:   MSI, NULLSOFT
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
    Write-Host "Save-AsPng ($SourceFile) to ($TargetFile) with a resolution of $($Resolution)x$($Resolution)" -ForegroundColor DarkCyan
    if ($SourceFile)
    {
        $iconBmp = [System.Drawing.Bitmap]::FromFile($SourceFile)
    }
    
    $newbmp = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $iconBmp, $Resolution, $Resolution
    write-host "Resize image to $($Resolution) x $($Resolution)"
    
    $newbmp.Save($TargetFile, "png")
    $newbmp.Dispose()
    try
    {
        Test-Path $TargetFile
        Write-Host "$TargetFile saved" -ForegroundColor Green
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
			
			#Create a View object
			# SELECT `Message` FROM `Error` WHERE `Error` = 1715 
			# [Microsoft.Deployment.WindowsInstaller.View]
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
    foreach ($exe in $ExeShortCuts) #Create Icon From Shortcuts
    {                
        $ExeWithIcon = $exe
        if ($InstallerDetails.DisplayIcon.contains("\$exe"))
        {                    
            break #Handle it like it is the "suggested" icon
        }
    }
    if ($ExeWithIcon -ne "")
    {
        if ($InstallerType -eq "nullsoft")
        {
            $process = Start-Process -FilePath "$CustomExtractor" -ArgumentList @("e", "`"$InstallerFile`"", "-ir!$ExeWithIcon", "-o`"$($Global:TempDirectory)`"","-y") -WindowStyle Hidden -Wait -PassThru
        }
        if ($InstallerType -eq "msi")
        {
            Write-Host "Try to extract icon from `"$ExeWithIcon`"" -ForegroundColor Magenta
            $process = Start-Process -FilePath "$CustomExtractor" -ArgumentList @("x", "`"$InstallerFile`"", "$($Global:TempDirectory)\","$ExeWithIcon") -WindowStyle Hidden -Wait -PassThru
        }
        if ($process.ExitCode -eq 0)
        {
            $extractedfiles = Get-ChildItem -Path $($Global:TempDirectory) -Filter $ExeWithIcon -Recurse
            foreach ($file in $extractedfiles)
            {
                Move-Item -Path $file.FullName -Destination $($Global:TempDirectory) -Force
            }
            $InstallerDetails.TemporaryIconFile = ("$($Global:TempDirectory)\$ExeWithIcon").Replace(".exe",".ico")
            $InstallerDetails.FileDetectionVersion = ((Get-Item "$($Global:TempDirectory)\$ExeWithIcon").VersionInfo.FileVersion).Replace(", ", ".")
            if (!(Test-Path -Path "FileSystem::$($InstallerDetails.TemporaryIconFile)"))
            {
                try
                {
                    $nouptput = Save-WingetBridgeAppIcon -SourceFile "$($Global:TempDirectory)\$ExeWithIcon" -TargetIconFile $InstallerDetails.TemporaryIconFile
                }
                catch
                {
                    Write-Host "We were not able to extract an icon from `"$($Global:TempDirectory)\$ExeWithIcon`"" -ForegroundColor Yellow
                } #could be an executable that does not contain an icon
            }
            if (Test-Path -Path "FileSystem::$($InstallerDetails.TemporaryIconFile)")
            {
                $InstallerDetails.BestIconResolution = Get-BestIconResolution -SourceFile "$($Global:TempDirectory)\$ExeWithIcon"
            }
        } else
        {
            Write-Host "Failed to extract executable" -ForegroundColor Red
        }
    } else { Write-Host "no icon found" }
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
    [string]$CustomMsiLessSource
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
                    #$ShortcutTarget = $ShortcutTarget.Replace("[APPLICATIONFOLDER]", "")
                    #$ShortcutTarget = $ShortcutTarget.Replace("[INSTALLFOLDER]", "")
                    #$ShortcutTarget = $ShortcutTarget.Replace("[INSTALLDIR]", "")
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

            $InstallerDetails = Get-WingetBridgeIcon -InstallerDetails $InstallerDetails -InstallerType "msi" -CustomExtractor $CustomMsiLessSource -ExeShortCuts $ExeShortCuts
            return $InstallerDetails
        }
        else
        {
            Write-Host "Setup does not look like an msi-installer!" -ForegroundColor Red
        }
    }
}

function Get-InstallerDetailsForNullsoft
{
param (
    [string]$InstallerFile,
    [string]$Custom7zSource
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
        Write-Host "Verify version of 7-Zip [$Custom7zSource]" -ForegroundColor Magenta
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
            $installerDetails = Get-WingetBridgeIcon -InstallerDetails $InstallerDetails -InstallerType "nullsoft" -CustomExtractor $Custom7zSource -ExeShortCuts $ExeShortCuts
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
    [string]$TempDirectory
)
    try
    {
        Remove-Item -Path "$TempDirectory\*.ico" -Filter "*.ico" -Force
        Remove-Item -Path "$TempDirectory\*.exe" -Filter "*.exe" -Force
        Remove-Item -Path "$TempDirectory\*.nsi" -Filter "*.nsi" -Force
        Remove-Item -Path "$TempDirectory\*.nsi" -Filter "*.png" -Force
    } catch {}
}