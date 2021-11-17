###
# Author:          Paul Jezek
# ScriptVersion:   v1.0.6, Nov 17, 2021
# Description:     WingetBridge Factory for MEMCM
# Compatibility:   MSI, NULLSOFT, INNO
# Please visit:    https://github.com/endpointmanager/wingetbridge-factory  - to get the latest version
#######

#######
# CONFIGURATION STARTS HERE
###

    # Specify whatever you want to "synchronize" (2900+ packages are available) - You can search for package-IDs by using the [Start-WingetSearch]-cmdlet, or visit https://winstall.app
    $WingetPackageIdsToSync = @( "Microsoft.VisualStudioCode", "VideoLAN.VLC" )
    $ContentSource = "\\MEMCM01\Setups"                                                        # Content will be automatically downloaded to a new subdirectory like [ContentSource]\Publisher\Appname\Architecture\SomeInstaller.msi
    $CMApplicationFolder = ""                                                                  # Specify a folder in MEMCM where to move newly created applications (e.g. LAB:\WingetBridgeFactory") If you leave this empty, it will be stored in the root-folder    

    # Custom install-parameters (optional)
    $InstallParams = @{}
    $InstallParams['Microsoft.VisualStudioCode'] = "/VERYSILENT /NORESTART /MERGETASKS=!runcode"           #This is just an example (for Microsoft.VisualStudioCode)
    $InstallParams['Microsoft.VisualStudioCode.Insiders'] = "/VERYSILENT /NORESTART /MERGETASKS=!runcode"  #This is just an example (for Microsoft.VisualStudioCode.Insiders)

    # Custom uninstall-parameters (optional)
    $UninstallParams = @{}

    # "WingetBridge Factory" configuration
    $Global:ScriptPath = $null                                                                 # Does not need to be configured manually (by default, the root directory of this script will be use)
    $Global:TempDirectory = "[SCRIPTDIR]\tmp"                                                  # Temp-directory for this script, Warning: LessMSI does not support spaces in TempDirectory for unknown reasons
    $Global:7z15Beta_exe = "[SCRIPTDIR]\bin\7z1505\7z.exe"                                     # Please download "7-Zip v15.05 beta" from https://sourceforge.net/projects/sevenzip/files/7-Zip/15.05/ (Warning: It has to be version "15.05 beta", as newer versions do not decode NSIS!)
    $Global:LessMSI_exe = "[SCRIPTDIR]\bin\lessmsi\lessmsi.exe"                                # Please download "LessMSI" from https://github.com/activescott/lessmsi/releases
    $Global:InnoUnp_exe = "[SCRIPTDIR]\bin\innounp\innounp.exe"                                # Please download "innounp, the Inno Setup Unpacker" from http://innounp.sourceforge.net
    $Global:WIXToolset_InstallPath = "%ProgramFiles(x86)%\WiX Toolset v3.11"                   # Please download "Wix Toolset" (3.11 or later) from https://wixtoolset.org/releases/ and specify the installation path (e.g. C:\Program Files (x86)\WiX Toolset v3.11)
    $Global:SharedFactoryScript = "[SCRIPTDIR]\WingetBridge-Factory_Shared.psm1"               # WingetBridge Factory - Shared functions

    $Global:DownloadTimeout = 900                             # Timout when downloading installers (Some proxy servers need a lot of time to check the content)
    $Global:AllowMSIWrappers = $false                         # Sometimes a legacy-setup is better than a MSI wrapping a legacy-installer
    $Global:InstallerLocale = @("en-US", "de-DE", $null)      # Some installers do not provide a locale (asume it is english or multilange), and some will overwrite already downloaded setups from a different locale (having the same filename)
    $Global:MaxCacheAge = 30                                  # By default wingetbridge caches winget repo for one day, you can specify a custom value in minutes
    $Global:CleanupFiles = $true                              # Remove temporary files from WingetBridgeFactory TempDirectory as soon as we don not need them anymore
    $Global:UseCachedAppIcons = $true                         # If set to true, a cached *.ico (available in "[ContentSource]\{Publisher}\{WingetPackageId}.ico") will be used. If not available, we try to extract it from an installer
    $Global:CacheAppIcons = $true                             # If set to true, an extracted icon will be cached as "[ContentSource]\{Publisher}\{WingetPackageId}.ico"
    $Global:VerboseMessages = $true                           # Use VerboseMessages (for Troubleshooting in WingetBridge Factory)

    # Site configuration
    $SiteCode = "LAB"                                         # Site code
    $ProviderMachineName = "MEMCM01"                          # SMS Provider machine name
    $initParams = @{}

    #If you prefer the "offline installation" of endpointmanager.wingetbridge: Please uncomment the next line and specify the path where you extracted the Release.zip of the module
    #Import-Module C:\YourData\ExtractedReleaseZip\endpointmanager.wingetbridge\x.x.x.x\endpointmanager.wingetbridge.psd1

###
# CONFIGURATION ENDS HERE
#######

function New-CMWingetBridgePackage #for MEMCM Script
{
param (
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Use Start-WingetSearch to get the required package",ValueFromPipeline=$true)]
    [endpointmanager.wingetbridge.WingetPackage]$WingetBridgePackage,
    [string]$ContentSourceRootDir,
    [string]$CMAppFolder,
    [HashTable]$CustomInstallParams,
    [HashTable]$CustomUninstallParams
)
    try
    {        
        $CMApp = $null        

        if ($WingetBridgePackage -eq $null)
        {
            Write-Host "You need to specify a package-manifest using" -ForegroundColor Red
        }
        else
        {
            $WingetBridgePackage
            $manifest = $WingetBridgePackage.LatestVersion | Get-WingetManifest

            $AppName = "$($manifest.PackageName) v$($manifest.PackageVersion)"
            $AppVersion = "$($manifest.PackageVersion)"
            $AdminComment = "Created with WingetBridge Factory ($($manifest.PackageIdentifier))"
            $FriendlyPublisherName = Get-PathFriendlyName -Name $manifest.Publisher

   
            if ($OperatingSystemCondition -eq $null)
            {
                $OperatingSystemCondition = Get-CMGlobalCondition | where {if ($_.CI_UniqueID -eq "GLOBAL/OperatingSystem"){ return $_;}} #can take long (but does not care about localization)
            }

            $SupportedInstallerFound = $false
            $AppCreationIsRequired = $true

            $CMApp = Get-CMApplication -Name "$AppName"
            if ($CMApp -ne $null)
            {
                Write-Host "$AppName already exists" -ForegroundColor Green
                $AppCreationIsRequired = $false                
            }
            else
            {
                Write-Host "$AppName does not exist yet. Let's create it ..." -ForegroundColor Yellow
                $keywords = Get-WingetBridgeKeywordsAsString -Tags $manifest.Tags -maxlength 64

                if ($manifest.Description -eq $null)
                {
                    $AppDescription = $($manifest.ShortDescription)
                }
                else
                {
                    $AppDescription = $($manifest.Description)
                }
        
                $PublisherRootDir = "$ContentSourceRootDir\$FriendlyPublisherName"
                $PackageIdRootDir = "$PublisherRootDir\$($AppName.Trim("."))"
                $IconResource = "$PublisherRootDir\$($manifest.PackageIdentifier.Trim(".")).ico" #Pre-cached Icon

                $x64InstallerFound = $false
                $x86InstallerFound = $false

                $IconAttached = $false
                $ExpectedScope = "machine"
                                
                $CreatedDTs = @()
                $testrun = $true
                for ($i=1; $i -le 2; $i++) #Loop twice (one testrun, to detect supported Installers, then for the "magic")
                {                                                            
                    foreach ($installer in $manifest.Installers)
                    {
                        if (($testrun -eq $false) -and ($AppCreationIsRequired -eq $true) -and ($SupportedInstallerFound -eq $true)) #Create app if supported installer found ...
                        {
                            Write-Host "Supported installer found, create Application" -ForegroundColor Green
                            $CMAppCreation = New-CMApplication -Name "$AppName" -Keyword $keywords -LocalizedName "$($manifest.PackageName)" -LocalizedDescription "$AppDescription" -Description "$AdminComment" -SoftwareVersion "$AppVersion" -Publisher $FriendlyPublisherName #Always start without icon (*.ico could fail if it contains layers with resolution higher than 512x512)
                            $CMApp = Get-CMApplication -Name "$AppName" #Verify creation                            

                            if ($CMApp -ne $null)
                            {
                                if (($CMAppFolder -ne $null) -and ($CMAppFolder -ne ""))
                                {
                                    Get-CMApplication -Name "$AppName" | Move-CMObject -FolderPath $CMAppFolder #Move to specified folder
                                }
                                $AppCreationIsRequired = $false

                                if (($Global:UseCachedAppIcons) -and (Test-Path -Path "FileSystem::$IconResource")) #Use a pre-cached icon
                                {                                                                
                                    $BestIconResolution = Get-BestIconResolution -SourceFile "$IconResource"
                                    if ($BestIconResolution -ne 0)
                                    {
                                        $PNGToImport = "$Global:TempDirectory\$([io.path]::GetFileName($IconResource).Replace(".ico",".png"))"
                                        Write-Host "Import AppIcon (from PNG): $PNGToImport"
                                        Save-AsPng -SourceFile $IconResource -TargetFile $PNGToImport -Resolution $BestIconResolution
                                        if (Test-Path -Path $PNGToImport)
                                        {
                                            $CMApp | Set-CMApplication -IconLocationFile $PNGToImport
                                            Write-Host "`"Pre-cached`" app-icon ($IconResource) was used" -ForegroundColor Gray
                                            $IconAttached = $true
                                        }
                                    }
                                }
                            }
                        }    

                        if (((($installer.Scope -eq $ExpectedScope) -or ($manifest.Scope -eq $ExpectedScope)) -or (($installer.Scope -eq $null) -and ($manifest.Scope -eq $null))) -and ($installer.Architecture -ne "arm64") -and (($manifest.InstallerLocale -in $Global:InstallerLocale) -or ($installer.InstallerLocale -in $Global:InstallerLocale))) #ignore mobile platform (arm64)
                        {
                            if ($testrun)
                            {
                                $installerType = $installer.InstallerType
                                if (($installerType -eq "") -or ($installerType -eq $null)) { $installerType = $manifest.InstallerType }

                                $installerLocale = $installer.InstallerLocale
                                if (($installerLocale -eq "") -or ($installerLocale -eq $null)) { $installerLocale = $manifest.InstallerLocale }
                    
                                $Scope = $installer.Scope
                                if (($Scope -eq "") -or ($Scope -eq $null)) { $Scope = $manifest.Scope }
                                Write-host "Found machine targeted installer ($installerType, $($installer.Architecture)  $($installer.InstallerLocale))" -ForegroundColor Magenta
                            }

                            if (($installerType -eq "nullsoft") -or ($installerType -eq "msi") -or ($installerType -eq "inno"))
                            {
                                $SupportedInstallerFound = $true
                                if ($installer.Architecture -eq "x64") { $x64InstallerFound = $true }
                                if ($installer.Architecture -eq "x86") { $x86InstallerFound = $true }
                                if ($testrun -eq $false)
                                {
                                    $SkipInstaller = $false
                                    if (($installerLocale -ne $null) -and ($installerLocale -ne ""))
                                    {
                                        $DeploymentTypeName = "$AppName $($installer.Architecture) ($Scope, $installerType, $installerLocale)"
                                    }
                                    else
                                    {
                                        $DeploymentTypeName = "$AppName $($installer.Architecture) ($Scope, $installerType)"
                                    }
                                    Write-Host -ForegroundColor Magenta "Download $($installer.InstallerUrl)"
                                    $ContentSource = "$PackageIdRootDir\$($installer.Architecture)"
                                    $NewContentFolder = New-Item -ItemType Directory -Path "FileSystem::$ContentSource" -Force
                                    $downloadedInstaller = Start-WingetInstallerDownload -PackageInstaller $installer -TargetDirectory $ContentSource -AcceptAgreements -Timeout $Global:DownloadTimeout                            

                                    if ($installerType -eq "inno")
                                    {
                                        $installerdetails = Get-InstallerDetailsForInno -InstallerFile "$ContentSource\$($downloadedInstaller.Filename)" -CustomInnoUnpSource $Global:InnoUnp_exe -ExtractIcon (!$IconAttached) #Extract Inno Compiler Instructions
                                    }
                                    if ($installerType -eq "nullsoft")
                                    {
                                        $installerdetails = Get-InstallerDetailsForNullsoft -InstallerFile "$ContentSource\$($downloadedInstaller.Filename)" -Custom7zSource $Global:7z15Beta_exe #Extract NullSoft Installer Instructions
                                    }
                                    if ($installerType -eq "msi")
                                    {
                                        $installerdetails = Get-InstallerDetailsForMSI -InstallerFile "$ContentSource\$($downloadedInstaller.Filename)" -CustomMsiLessSource $Global:LessMSI_exe #Extract MSI Informations
                                    }
                                    $installerdetails

                                    if ($installerdetails.DisplayName -ne "") {
                                        if ($installerdetails.DisplayName.ToLower().Contains($installer.Architecture.ToLower()))
                                        {
                                            $DeploymentTypeName = $installerdetails.DisplayName
                                        }
                                        else
                                        {
                                            $DeploymentTypeName = "$($installerdetails.DisplayName) ($($installer.Architecture))"
                                        }
                                    }
                            
                                    if ($installerType -eq "msi")
                                    {
                                        if ($installerdetails.FileCount -eq 0)
                                        {
                                            Write-Host "This is just a MSI-Wrapper including $($installerdetails.BinaryDataCount) binaries. (Please consider using a different InstallerType)" -ForegroundColor Red
                                            if (!($Global:AllowMSIWrappers)) { $SkipInstaller = $true }
                                        }
                                    }

                                    if (($installer.Architecture -ne "x64") -and ($installer.Architecture -ne "x86"))
                                    {
                                        Write-Host "Skip installer, with unexpected architecture" -ForegroundColor Red
                                        $SkipInstaller = $true
                                    }                                    

                                    #Check if we already created an equivalent DT
                                    if ($CreatedDTs -contains "$Scope_$($installer.Architecture)_$installerLocale")
                                    {
                                        Write-Host "Skip Installer, as we already created an equivalent DeploymentType (same Scope, Architecture and Localization)" -ForegroundColor Yellow
                                        $SkipInstaller = $true
                                    }

                                    if (!($SkipInstaller))
                                    {
                                        if (($x86InstallerFound -eq $true) -and ($x64InstallerFound -eq $true))
                                        {
                                            $SupportedPlatform = @()                   
                                            if ($installer.Architecture -eq "x64")
                                            {
                                                $SupportedPlatform += Get-CMConfigurationPlatform -Name "All*Windows 10*(64-bit)*" | Where-Object PlatformType -eq 1
                                                $SupportedPlatform += Get-CMConfigurationPlatform -Name "All*Windows 11*(64-bit)*" | Where-Object PlatformType -eq 1
                                            }
                                            if ($installer.Architecture -eq "x86")
                                            {
                                                $SupportedPlatform += Get-CMConfigurationPlatform -Name "All*Windows 10*(32-bit)*" | Where-Object PlatformType -eq 1
                                                $SupportedPlatform += Get-CMConfigurationPlatform -Name "All*Windows 11*(32-bit)*" | Where-Object PlatformType -eq 1
                                            }
                                            $OSRequirements = $OperatingSystemCondition | New-CMRequirementRuleOperatingSystemValue -RuleOperator OneOf -Platform $SupportedPlatform
                                            Write-Host "Created OperatingSystem-RequirementRule (for $($installer.Architecture))"
                                        }
                                        else
                                        {
                                            $OSRequirements = @()
                                        }
                                                                            
                                        if (($IconAttached -eq $false) -and ($installerdetails.TemporaryIconFile -ne "")) #Set the AppIcon
                                        {
                                            if ($Global:CacheAppIcons)
                                            {
                                                if ((Test-Path -Path "FileSystem::$($installerdetails.TemporaryIconFile)") -and (!(Test-Path -Path "FileSystem::$IconResource"))) #Save the Icon we extracted
                                                {                                                
                                                    Copy-Item "FileSystem::$($installerdetails.TemporaryIconFile)" -Destination "FileSystem::$IconResource"
                                                }
                                            }
                                            if ($installerdetails.BestIconResolution -gt 0)
                                            {
                                                $PNGToImport = $installerdetails.TemporaryIconFile.Replace(".ico",".png")
                                                Write-Host "Import AppIcon (from PNG): $PNGToImport"
                                                Save-AsPng -SourceFile $IconResource -TargetFile $PNGToImport -Resolution $installerdetails.BestIconResolution
                                                Get-CMApplication -Name $AppName | Set-CMApplication -IconLocationFile $PNGToImport
                                                $IconAttached = $true
                                            }
                                        }                            

                                        if ($Global:CleanupFiles)
                                        {
                                            Remove-WingetBridgeFactoryTempFiles -TempDirectory $Global:TempDirectory -FullCleanup $false #Now that we have the icon, we can delete the temporary files                                        
                                        }

                                        ###
                                        # Create MSI-DeploymentType (for MSI)
                                        ##
                                        if ($installerType -eq "msi")
                                        {
                                            if ($CustomInstallParams[$WingetBridgePackage.PackageId] -eq $null)
                                            {
                                                $InstallProgram = "msiexec /I `"$($downloadedInstaller.Filename)`" /Q"
                                            }
                                            else
                                            {
                                                $InstallProgram = "msiexec /I `"$($downloadedInstaller.Filename)`" $($CustomInstallParams[$WingetBridgePackage.PackageId])" #use custom install parameters
                                            }
                                            if ($CustomUninstallParams[$WingetBridgePackage.PackageId] -eq $null)
                                            {
                                                $UninstallProgram = "msiexec /X $($installerdetails.ProductCode) /Q"
                                            }
                                            else
                                            {
                                                $UninstallProgram = "msiexec /X $($installerdetails.ProductCode) $($CustomUninstallParams[$WingetBridgePackage.PackageId])" #use custom uninstall parameters
                                            }
                                            $CMDeploymentType = Add-CMMsiDeploymentType -ApplicationName "$AppName" -DeploymentTypeName $DeploymentTypeName -ContentLocation "$ContentSource\$($downloadedInstaller.Filename)" -LogonRequirementType WhereOrNotUserLoggedOn -UninstallOption "NoneRequired" -InstallationBehaviorType InstallForSystem -UninstallCommand $UninstallProgram -AddRequirement $OSRequirements -UserInteractionMode Hidden
                                        }

                                        ###
                                        # Create Script-DeploymentType with Registry-or File-Detection (for NULLSOFT and INNO)
                                        ##
                                        if (($installerType -eq "nullsoft") -or ($installerType -eq "inno"))
                                        {
                                            if ($CustomInstallParams[$WingetBridgePackage.PackageId] -eq $null)
                                            {
                                                if ($installerType -eq "nullsoft") {
                                                    $InstallProgram = "`"$($downloadedInstaller.Filename)`" /S /allusers" #nullsoft installer (machine)
                                                }
                                                if ($installerType -eq "inno") {
                                                    $InstallProgram = "`"$($downloadedInstaller.Filename)`" /VERYSILENT /ALLUSERS /NORESTART" #inno installer (machine)
                                                }
                                            }
                                            else
                                            {
                                                $InstallProgram = "`"$($downloadedInstaller.Filename)`" $($CustomInstallParams[$WingetBridgePackage.PackageId])" #use custom install parameters
                                            }
                                            $RequiresWow6432 = $false                                            
                                            if ($($installerdetails.UninstallString) -ne "")
                                            {
                                                if ($CustomUninstallParams[$WingetBridgePackage.PackageId] -eq $null)
                                                {
                                                    if ($installerType -eq "nullsoft") {
                                                        $UninstallProgram = "`"$($installerdetails.UninstallString)`" /S /allusers _?=$($installerdetails.InstallDir)" #do not use quotes after _? (even if path contains spaces)
                                                        $UninstallProgram = $UninstallProgram.Replace("`$INSTDIR", $($installerdetails.InstallDir))

                                                        $UninstallProgram = $UninstallProgram.Replace("`$PROGRAMFILES64", "%ProgramFiles%")
                                                        $UninstallProgram = $UninstallProgram.Replace("`$PROGRAMFILES32", "%ProgramFiles(x86)%")
                                                        if ($UninstallProgram.Contains("`$PROGRAMFILES"))
                                                        {
                                                            if ($installer.Architecture -eq "x64") #Targeted for x64, but contains 32bit App-Directory
                                                            {
                                                                $RequiresWow6432 = $true #Used later for Registry-Entries (Uninstaller, which might be in Wow6432Node)
                                                                $UninstallProgram = $UninstallProgram.Replace("`$PROGRAMFILES", "%ProgramFiles(x86)%")
                                                            }
                                                            else #Targeted for x86
                                                            {                                
                                                                $UninstallProgram = $UninstallProgram.Replace("`$PROGRAMFILES", "%ProgramFiles(x86)%") #Does that work on a x86(??)
                                                            }
                                                        }
                                                    }
                                                    if ($installerType -eq "inno") {
                                                        $UninstallProgram = "`"$($installerdetails.UninstallString)`" /VERYSILENT /NORESTART"
                                                    }
                                                }
                                                else
                                                {
                                                    $UninstallProgram = "`"$($installerdetails.UninstallString)`" $($CustomUninstallParams[$WingetBridgePackage.PackageId])" #use custom uninstall parameters
                                                }
                                            }
                                            else #Could not detect a usefull uninstaller (needs to be specified manually)
                                            {
                                                $UninstallProgram = "`"$($installerdetails.UninstallString)`" $($CustomUninstallParams[$WingetBridgePackage.PackageId])" #use custom uninstall parameters
                                            }

                                            if (($installerType -eq "inno") -and ($($installerdetails.Is64bit) -eq $false)) { $RequiresWow6432 = $true }

                                            $CMDeploymentType = Add-CMScriptDeploymentType -ScriptLanguage VBScript -LogonRequirementType WhetherOrNotUserLoggedOn -ScriptText "option explicit" -ApplicationName "$AppName" -DeploymentTypeName $DeploymentTypeName -ContentLocation "$ContentSource" -UninstallOption "NoneRequired" -InstallationBehaviorType InstallForSystem -InstallCommand "$InstallProgram" -UninstallCommand $UninstallProgram                                           

                                            #Get Detection-Method from current Installer
                                            $DetectionClauses = @()
                                            $DetectionRegKey = $installerdetails.UninstallRegKey                            
                                            if (($DetectionRegKey -eq "") -or ($DetectionRegKey -eq $null))
                                            {
                                                Write-Host "Installer seems to be `"Registry-Free`", try to use file-detection..."
                                                if (($installerdetails.FileDetection -ne "") -and ($installerdetails.FileDetection -ne $null))
                                                {
                                                    $FileToDetect = [io.path]::GetFileName($installerdetails.FileDetection)
                                                    $FilePathToDetect = [io.path]::GetDirectoryName($installerdetails.FileDetection)                                
                                                    $ExpectedFileVersion = $($installerdetails.FileDetectionVersion).Replace(",","")
                                                    Write-Host "Expected FileVersion is $ExpectedFileVersion"
                                                    $FileDetection = New-CMDetectionClauseFile -FileName $FileToDetect -Is64Bit:$true -Path $FilePathToDetect -ExpressionOperator IsEquals -PropertyType Version -ExpectedValue $ExpectedFileVersion -Value
                                                    $DetectionClauses += $FileDetection
                                                }
                                                else
                                                {
                                                    Write-Host "We did not found a useable file- oder registry-detection in this installer. Requires manual configuration!" -ForegroundColor Yellow
                                                }
                                            }
                                            else
                                            {
                                                #Create Detection Method
                                                $RegistryDetection = $null
                                                $RegistryDetection = New-CMDetectionClauseRegistryKeyValue -Is64Bit:(($installer.Architecture -eq "x64") -and ($RequiresWow6432 -eq $false)) -Hive LocalMachine -KeyName $DetectionRegKey -PropertyType "String" -ValueName "DisplayVersion" -ExpressionOperator IsEquals -Value -ExpectedValue $installerdetails.DisplayVersion

                                                #Create alternative Detection Method
                                                if ($installer.Architecture -eq "x64")
                                                {
                                                    $Wow6432RegistryDetection = $null
                                                    $Wow6432RegistryDetection = New-CMDetectionClauseRegistryKeyValue -Is64Bit:$false -Hive LocalMachine -KeyName $DetectionRegKey -PropertyType "String" -ValueName "DisplayVersion" -ExpressionOperator IsEquals -Value -ExpectedValue $installerdetails.DisplayVersion
                                                }
                                
                                                $DetectionClauses += $RegistryDetection
                                                if ($installer.Architecture -eq "x64") { $DetectionClauses += $Wow6432RegistryDetection } #as we are not able to fully decode nullsoft-script, asume it could be registered in Wow6432Node                    
                                            }


                                            if ($DetectionClauses.count -gt 1){
                                                $DetectionClauseConnector = @()
                                                For ($i=1; $i -lt $DetectionClauses.length; $i++) {
                                                    $logicName = $DetectionClauses[$i].Setting.LogicalName
                                                    $DetectionClauseConnector += @{"LogicalName"= $logicName;Connector="OR"}
                                                }                        
                                            }

                                            Write-Host " Created Detection Method for Key: $DetectionRegKey Property: $DetectionRegKeyValueName Value: $SelectedVersion" -ForegroundColor Green
                                            Write-Host " Adding Detection Method to App" -ForegroundColor Green

                                            $SetDetection = $null
                                            if ($DetectionClauses.count -gt 1){
                                                #Include an $DetectionClauseConnector
                                                $SetDetection = Set-CMScriptDeploymentType -ApplicationName $AppName -DeploymentTypeName $DeploymentTypeName -UserInteractionMode Hidden -AddDetectionClause $DetectionClauses -DetectionClauseConnector $DetectionClauseConnector -AddRequirement $OSRequirements
                                            }
                                            else
                                            {
                                                $SetDetection = Set-CMScriptDeploymentType -ApplicationName $AppName -DeploymentTypeName $DeploymentTypeName -UserInteractionMode Hidden -AddDetectionClause $DetectionClauses -AddRequirement $OSRequirements
                                            }
                                        } #NULLSOFT DEPLOYMENT TYPE
                                        $CreatedDTs += "$Scope_$($installer.Architecture)_$installerLocale"  #Store what we already created in this package
                                    } #SkipInstaller
                                }
                            }
                        }                        
                    } ## foreach installer
                    $testrun = $false #next loop wont be a testrun
                } ## Loop twice (inside the same installer)
                if ($SupportedInstallerFound -ne $true)
                {
                    Write-Host "Package found, but no supported installer" -ForegroundColor Yellow
                }
            }
        }
        return $CMApp
    }
    catch
    {
        $formatstring = "`n{0} : {1}`n{2}`n" +
                        "    + CategoryInfo          : {3}`n" +
                        "    + FullyQualifiedErrorId : {4}`n" +
                        "    + ErrorMessage          : {5}`n"

        $fields = $_.InvocationInfo.MyCommand.Name,
                  $_.ErrorDetails.Message,
                  $_.InvocationInfo.PositionMessage,
                  $_.CategoryInfo.ToString(),
                  $_.FullyQualifiedErrorId,
                  $_
    
        Write-Host "$($formatstring -f $fields)" -ForegroundColor Red
    }
}

#######
## MAIN CODE STARTS HERE
#

    # Handle [SCRIPTDIR] in configuration ...
    if ($Global:ScriptPath -eq $null)
    {
        try
        {
            $ScriptDir = $(Split-Path $Script:MyInvocation.MyCommand.Path)
        } catch {}
        if ($ScriptDir -eq $null) { $ScriptDir = (Get-Location -PSProvider FileSystem).Path }
    } else { $ScriptDir = $Global:ScriptPath }

    Write-Host "[SCRIPTDIR] is set to: `"$ScriptDir`"" -ForegroundColor Gray

    # Load ConfigMgr Module
    if((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
    }

    # Connect to the site's drive if it is not already present
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams

    #Replace [SCRIPTDIR] in our variables of script-dependencies
    $Global:TempDirectory = $Global:TempDirectory.Replace("[SCRIPTDIR]", "$ScriptDir")
    $Global:7z15Beta_exe = $Global:7z15Beta_exe.Replace("[SCRIPTDIR]", "$ScriptDir")
    $Global:LessMSI_exe = $Global:LessMSI_exe.Replace("[SCRIPTDIR]", "$ScriptDir")
    $Global:InnoUnp_exe = $Global:InnoUnp_exe.Replace("[SCRIPTDIR]", "$ScriptDir")    
    $Global:SharedFactoryScript = $Global:SharedFactoryScript.Replace("[SCRIPTDIR]", "$ScriptDir")

    $Global:WIXToolset_InstallPath = $Global:WIXToolset_InstallPath.Replace("[SCRIPTDIR]", "$ScriptDir")
    $Global:WIXToolset_InstallPath = $Global:WIXToolset_InstallPath.Replace("%ProgramFiles(x86)%", [Environment]::GetEnvironmentVariable("ProgramFiles(x86)"))
    $Global:WIXToolset_InstallPath = $Global:WIXToolset_InstallPath.Replace("%ProgramFiles%", [Environment]::GetEnvironmentVariable("ProgramFiles"))
    $Global:WIXToolset_InstallPath = $Global:WIXToolset_InstallPath.TrimEnd("\")
    $Global:WindowsInstaller_dll = "$($Global:WIXToolset_InstallPath)\bin\Microsoft.Deployment.WindowsInstaller.dll"
    $Global:WIX_Dark_exe = "$($Global:WIXToolset_InstallPath)\bin\dark.exe"

    #Verify dependencies of WingetBridge Factory ...
    if ((!(Test-Path -Path "FileSystem::$Global:WindowsInstaller_dll")) -or (!(Test-Path -Path "FileSystem::$Global:WIX_Dark_exe"))) { Write-Host "ERROR: Please install WIX Toolset before using WingetBridge Factory!" -ForegroundColor Red
        break
    }
    if (!(Test-Path -Path "FileSystem::$Global:InnoUnp_exe")) { Write-Host "ERROR: Please install Inno Setup Unpacker (innounp, from {http://innounp.sourceforge.net/}) before using WingetBridge Factory!" -ForegroundColor Red
        break
    }
    if (!(Test-Path -Path "FileSystem::$Global:LessMSI_exe")) { Write-Host "ERROR: Please install MSI Unpacker (lessmsi, {from https://github.com/activescott/lessmsi/releases}) before using WingetBridge Factory!" -ForegroundColor Red
        break
    }
    if (!(Test-Path -Path "FileSystem::$Global:7z15Beta_exe")) { Write-Host "ERROR: Please install the correct version of 7-Zip (15.05 beta, {from https://sourceforge.net/projects/sevenzip/files/7-Zip/15.05/}) before using WingetBridge Factory!" -ForegroundColor Red
        break
    }
	$getwgm = Get-Command Get-WingetManifest -ErrorAction Ignore
	if (!(($getwgm -ne $null) -and (($getwgm).Version -ige "1.2.0.0"))) { Write-Host "ERROR: Please install the latest version of WingetBridge (1.2.0.0 or later {from https://github.com/endpointmanager/wingetbridge-powershell}) before using WingetBridge Factory!" -ForegroundColor Red
		break
	}

    #Create the required subdirectories
    $nf = New-Item -ItemType Directory -Path "FileSystem::$ScriptDir\bin" -Force
    $nf = New-Item -ItemType Directory -Path "FileSystem::$ScriptDir\tmp" -Force

    # Load the shared functions of WingetBridge Factory
    Import-Module "$SharedFactoryScript"

    # Lets begin with the magic...
    foreach ($IdToSync in $WingetPackageIdsToSync)
    {
        Write-Host "Search for `"$IdToSync`" in winget repository ..."
        $CreatedPackage = Start-WingetSearch -SearchById $IdToSync -MaxCacheAge $Global:MaxCacheAge | New-CMWingetBridgePackage -ContentSourceRootDir $ContentSource -CMAppFolder $CMApplicationFolder -CustomInstallParams $InstallParams -CustomUninstallParams $UninstallParams
        Write-Host ""
    }
    if ($Global:CleanupFiles)
    {
        Remove-WingetBridgeFactoryTempFiles -TempDirectory $Global:TempDirectory -FullCleanup $true
    }  
    Write-Host "We are done! Thank you for using WingetBridge Factory!" -ForegroundColor Green

#
## MAIN CODE ENDS HERE
#######