# KMS_VL_ALL_AIO

## Table of Content
|Thing                               |
|------------------------------------|
|[Introduction](#introduction)       |
|[Supported Products](#supported)    |
|[Unsupported Products](#unsupported)|
|[How to Install](#install)          |
|[How to Use](#use)                  |

<h2 id="introduction">Smart Activation Script</h2>

KMS_VL_ALL_AIO is an All-in-One activation script for Microsoft Windows, Microsoft Office 2013 - 2021, and Microsoft Office 2010.

By design, the activation period lasts up to 180 days at most. This script also provides an auto renewal for continuous reactivation.

<h2 id="supported">Supported products </h2>

The list of supported products:
- Windows 11, 11 ARM64
- Windows 10, 10 ARM64
- Windows 8.1
- Windows 8
- Windows 8, 8.1, 10, 11 N Editions
- Windows 7 (Enterprise /N/E, Professional /N/E, Embedded POSReady/Thin PC)
- Windows Server 2022, 2019, 2016
- Windows Server 2012 & 2012 R2
- Windows Server 2008 R2
- Microsoft Office 2010, 2013, 2016, 2019, and 2021

*These editions are only activatable for 45 days max.*
- Windows 10, 11 Home edition variants
- Windows 8.1 Core edition variants, Pro with Media Center, Pro Student

*These editions are only activatable for 30 days max.*
- Windows 8 Core edition variants, Pro with Media Center

<h2 id="unsupported">Unsupported Products</h2>

The list of unsupported products:
- Office MSI Retail 2010, 2013 & Office 2010 C2R Retail
- Office UWP Universal Windows Platform (Windows 10/11 Apps)
- Windows Evaluation Editions
- Windows 7 (Starter, HomBasic, HomePremium, Ultimate)
- Windows 10 (Cloud "S", IoT Enterprise, IoS EnterpriseS, Professional SingleLanguage, Professional China, etc.)
- Windows 11 (IoT Enterprise, Professional SingleLanguage, Professional China, etc.)
- Windows Server (Azure Stack HCL, Server Foundation, Storage Server, Home Server 2011, etc.)

<h2 id="install">How to install</h2 >

*Before using, make sure any other KMS solutions are removed from the system first.*

- Download the latest activation script from the release page
- Extract the KMS_VL_ALL_AIO-xx.cmd with the password being whatever year that release is from. For example: Version 49 was released in 2023, the password is ```2023```
- [Disable Windows Defender](https://youtu.be/UKu6qtc534A) or allow KMS_VL_ALL_AIO.cmd on the device
- Run the KMS_VL_ALL_AIO.cmd script and follow the instructions.

<h2 id="use">How to use </h2>

### Activation Methods

Within this script, there are 3 activation methods:
- Manual: The activation script is executed and leaves no traves of KMS emulator on the system. This method does not reactivate products once they expire after the 180 days duration.
- Auto-Renewal: This is the **recommended** mode. This method installs a dll file that will automatically renew the license every 180 days. Newly installed Volume Office Products will be auto activated with this mode, although if you want to convert and activate Office C2R, renewing activation or activating new product, you will need to run ```Activate [Auto Renewal Mode]``` from the script menu again.
- External Mode: Standalone mode, where you activate against trusted external KMS server, without using the local KMS emulator. The external server can be an web address, or IP address.

To switch to Auto Renewal Mode, press <kbd>2</kbd> in the activation script to install the dll file.
To switch turn off Auto Renewal Mode, press <kbd>3</kbd>.

**Uninstall Completely** removes the auto renewal dll, leftover from a prematurely terminated script, clears the External mode server registration and traces, clears KMS cache, as well as OEM folder project.

### Configuration Mode

- Debug Mode: With this mode enabled, instead of outputing to the console directly, it will output to a log file where it can be read later for troubleshooting.
- Process Windows/Office: This mode is **on** by default. With this mode on, the activation script will try and activate them when ran. *Turning this mode OFF is not very effective if products are already Volume (GLVK Installed) because the system itself may try to reach and KMS activate products, especially on Windows 8 onward.*
- Convert Office C2R-R2V: Converts Office C2R Retail to Volume (unless Retail products already activated). This is **on** by default.
- Override Office C2R vNext: Overrides Office C2R vNext license (subscription or lifetime). This is **on** by default.
- Skip Windows KMS38: If KMS 2038 is deteced, the script skips Windows Activation. This mode is **on** by default.

### Miscellaneous Options

- Check Activation Status {vbs}: Shows the activation expiration date for Windows. Office 2010 ospp.vbs shows very little info.
- Check Activation Status {wmi}: Shows activation expiration date for all products as well as more info on Office 2010. Shows the status of Office UWP apps as well as more info (SKU ID, key channel).
- Create \$OEM\$ Folder: Creates needed folder structure and scripts to use during Windows installation to preactive the system.

[![Release downloads](https://img.shields.io/github/downloads/abbodi1406/KMS_VL_ALL_AIO/total.svg)](https://GitHub.com/abbodi1406/KMS_VL_ALL_AIO/releases/)
