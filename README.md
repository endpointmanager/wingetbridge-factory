![Image of wingetbridge-factory](https://repository-images.githubusercontent.com/427951913/6f318b11-7464-414a-a4db-09a1958fc808)
[![Twitter URL](https://img.shields.io/twitter/url/https/twitter.com/PaulJezek.svg?style=social&label=Follow%20%40Paul%20Jezek%20%23wingetbridge)](https://twitter.com/PaulJezek)
# wingetbridge-factory

This repository contains Powershell-Scripts based on [WingetBridge-PSModule](https://github.com/endpointmanager/wingetbridge-powershell) to automatically provide and maintain applications in various deployment tools by using the Windows Package Manager Repository.  

WingetBridge Factory creates new software packages in your deployment tool, based on a list of PackageID's that you specified in this script. You can schedule the script (e.g. on a daily basis), so that new versions of software packages will be created as soon as new versions become available in the winget repository.

## Community-driven Project

While there is only one example (for MEMCM, formerly ConfigMgr) in this repository, I'm very mind-opened in supporting a community-driven collaboration to provide more examples on how to use WingetBridge for various deployment tools like Microsoft Intune, MDT, DCM or VMWare Airwatch.  

If you are interested in contributing, please submit a pull request (PR) or contact me on [twitter](https://twitter.com/PaulJezek).

## Compatibility

At the moment the windows package manager repository contains 2900+ software packages.  

Each of them contains one or more installers (for various architectures, operating systems or languages) by using different installer technologies like MSI, NULLSOFT, INNO, Legacy-Setups (EXE), MSIX, APPX and APPXBUNDLE.  

Unfortunately, the winget repository does not provide a reliable detection-method (which is required for most software deployment tools on the market) for all installer types. It also does not provide any app-icons (e.g. to be used in a self-service-portal like Softwarecenter). WingetBridge Factory tries to extracts that information from the downloaded setup itself.
This method will not always deliver a 100% reliable detection-method and needs to get verified manually after the package was created.  

At the moment, the available scripts are intended to be used with MEMCM (formerly ConfigMgr) and only supports MSI and NULLSOFT installers. However, the MEMCM-script can be extended or even adopted for a different software deployment tool, depending on your own scripting-skills.

## Risk of damage :warning:

I highly recommend not using WingetBridge in a production environment without validating the downloaded installers before you deploy it. (e.g. by validating the certificate of a signed installer)

# WingetBridge Factory for MEMCM

The MEMCM version of the script does not respect the "MinimumOSVersion" information provided in the winget-manifest.  
The function "New-CMWingetBridgePackage" is the main-function of the MEMCM asumes that a downloaded setup is compatible with the same architecture of Windows 10 and Windows 11 and set up a requirement for it.

## Planned features

* Support for INNO Setups will be available within the next few weeks.
* (Optional) E-Mail notification when a new software package was created
* Automatically create supersedence (MEMCM)

Follow me on twitter and github to get notified about updates.

# Initial Configuration
In the initial setup, you specify a list of package-IDs you want to "synchronize" the latest available version into your software deployment tool.

There are multiple ways to find PackageIDs within the winget repository:
* By using the WingetBridge-PSModule(https://github.com/endpointmanager/wingetbridge-powershell)
* By using the winget-cli(https://github.com/microsoft/winget-cli)
* winstall.app, Online-WebUI(https://winstall.app/)

### # CONFIGURATION STARTS HERE
This is where the configuration starts.

* $WingetPackageIdsToSync -> Here you specify the list of PackageIDs that you want to get automatically maintained by Wingetbridge Factory
* Other variables are described inside the script.

### # CONFIGURATION ENDS HERE
This is where the configuration ends.  

If you want to modify or extend the functionality of the script, I highly recommend to modify only lines below ## MAIN CODE STARTS HERE, so you can easily migrate to a newer version as soon as it will become available.

## Requirements

The scripts in this repository, depends on several open source projects:

* [WingetBridge-Powershell-Module](https://github.com/endpointmanager/wingetbridge-powershell)
* [7-Zip v15.05 beta**](https://sourceforge.net/projects/sevenzip/files/7-Zip/15.05/) (Required to extract and analyze NULLSOFT-Scripts, and to extract Icons)
* [WIX Toolset](https://wixtoolset.org/) (Required to analyze MSI Setups)
* [lessMSI](https://lessmsi.activescott.com/) (Required to extract Icons from big MSI-files)

> **newer versions of 7-Zip do not extract and decode the required NSIS (nullsoft installer script). Please use v15.05!

### Additional requirements for MEMCM:

* Configuration Manager Admin-Console needs to be installed (Script depends on ConfigurationManager.psd1)

# Available Scripts
WingetBridge-Factory_MEMCM.ps1  
WingetBridge-Factory_Shared.psm1 - Module required by all other scripts (contains shared functions)

## Credits :heart:

* The developers and contributors behind [Windows Package Manager](https://docs.microsoft.com/en-us/windows/package-manager/).
* The developers and contributors behind [WIX Toolset](https://wixtoolset.org/)
* Igor Pavlov and the contributors of [7-Zip](https://www.7-zip.org/)
* Scott Willeke and the contributors of [lessmsi](https://lessmsi.activescott.com/)
* Everyone who is willing to use wingetbridge, so software vendors see how important it is to maintain their packages through the winget-repository.

## License (does not include 3rd party software, which can be downloaded using this script)

See [LICENSE](LICENSE) file for licence rights and limitations (MIT)

The author of this module is not responsible for, nor does it grant any licenses to, third-party packages.