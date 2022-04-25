[![Build Status](https://img.shields.io/github/workflow/status/rcmaehl/MSEdgeRedirect/mser)](https://github.com/rcmaehl/MSEdgeRedirect/actions?query=workflow%3Amser)
[![Download](https://img.shields.io/github/v/release/rcmaehl/MSEdgeRedirect)](https://github.com/rcmaehl/MSEdgeRedirect/releases/latest/)
[![Download count](https://img.shields.io/github/downloads/rcmaehl/MSEdgeRedirect/total?label=Downloads)](https://github.com/rcmaehl/MSEdgeRedirect/releases/latest/)
[![Chocolatey download count](https://img.shields.io/chocolatey/dt/msedgeredirect?label=Chocolatey+downloads)](https://chocolatey.org/packages/msedgeredirect)
[![Ko-fi](https://img.shields.io/badge/Support%20me%20on-Ko--fi-FF5E5B.svg?logo=ko-fi)](https://ko-fi.com/rcmaehl)
[![PayPal](https://img.shields.io/badge/Donate%20on-PayPal-00457C.svg?logo=paypal)](https://paypal.me/rhsky)
[![Join the Discord chat](https://img.shields.io/badge/Discord-chat-7289da.svg?&logo=discord)](https://discord.gg/uBnBcBx)

# MSEdgeRedirect
A Tool to Redirect News, Search, Widgets, Weather and More to Your Default Browser

This tool filters and passes the command line arguments of Microsoft Edge processes into your default browser instead of hooking into the `microsoft-edge:` handler, this should provide resiliency against future changes. Additionally, an Image File Execution Options mode is available to operate similarly to the Old EdgeDeflector. Additional modes are planned for future versions.

No Default App walkthrough or other steps, just set and forget.

> ℹ️ If you're on older Windows builds, check out: https://github.com/da2x/EdgeDeflector

> :warning: PLEASE NOTE: MSEdgeRedirect is still BETA. Changes are to be expected, and performance to be improved.

## Download

[Download latest stable release](https://github.com/rcmaehl/MSEdgeRedirect/releases/latest/download/MSEdgeRedirect.exe)

[Download latest testing release](https://nightly.link/rcmaehl/MSEdgeRedirect/workflows/MSER/main/mser.zip)\
**Keep in mind that you will have to update testing releases manually**

### Package Managers

[![image](https://user-images.githubusercontent.com/716581/159123573-58e5ccba-5c82-46ec-adcc-08b897284a6d.png)](https://github.com/microsoft/winget-pkgs/tree/master/manifests/r/rcmaehl/MSEdgeRedirect)  
```powershell
winget install MSEdgeRedirect
```

[![image](https://user-images.githubusercontent.com/716581/159197666-761d9b5e-18f6-427c-bae7-2cc6bd348b9a.png)](https://community.chocolatey.org/packages/msedgeredirect/)
```powershell
choco install msedgeredirect
```

### System Requirements
 |Minimum Requirements|Recommended
----|----|----
OS|Windows 8.1|Latest Windows 11 Build
CPU|32-bit Single Core|64-bit Dual Core or Higher
RAM (Memory)|40MB Free|100MB Free
Disk (Storage)|5MB Free|100MB Free

## Program Comparisons
 |[Edge Deflector](https://github.com/da2x/EdgeDeflector)|[ChrEdgeFkOff](https://github.com/AveYo/fox/blob/main/ChrEdgeFkOff.cmd)|[NoMoreEdge](https://github.com/HarshalKudale/NoMoreEdge)|[Search Deflector](https://github.com/spikespaz/search-deflector)|[MSEdgeRedirect](https://github.com/rcmaehl/MSEdgeRedirect/)
----|----|----|----|----|----
Redirection Modes|URI Handler<br/><br/>|IFEO<br/><br/>|IFEO<br/><br/>|URI Handler<br/><br/>|URI Detection<sup>*</sup><br/>or IFEO<sup>†</sup>
Redirects Search|☑|☑|☑|☑|☑
Installs without Admin|☑| | | |☑<sup>*</sup>
Windows 11 Support| |☑|☑| |☑
Windows 10 21H2+ Support| |☑|☑| |☑
Windows on ARM Support| |☑|☑| |
Installs System Wide| |☑|☑|☑, Optionally|☑<sup>†</sup>
Can be silently Deployed| |☑|☑| |☑
Update Checker Module| | |☑|☑|☑
Search Engine Customizations| | |☑, 8|☑, 14|☑, 9
Search Engine Shortcut Support| | |☑, DDG !Bang| | 
Weather Forecast Customizations| | | | |☑, 6
Customizable Edge Support| | | | |☑<sup>†</sup>
Can be used Portably (USB)| | | | |☑<sup>‡</sup>
Supports GPO Registry Keys| | | | |☑


<sub><sup>\* When using Service Mode, † When using Active Mode, ‡ When using /portable flag, uses Service Mode</sub></sup>

## Compiling

1. Download and run "AutoIt Full Installation" from [official website](https://www.autoitscript.com/site/autoit/downloads). 
1. Get the source code either by [downloading zip](https://github.com/rcmaehl/MSEdgeRedirect/archive/main.zip) or do `git clone https://github.com/rcmaehl/MSEdgeRedirect`.
1. Right click on `MSEdgeRedirect.au3` in the MSEdgeRedirect directory and select "Compile Script (x64) (or x86 if you have 32 bit Windows install).
1. This will create MSEdgeRedirect.exe in the same directory.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for rules of coding and pull requests.

## License

MSEdgeRedirect is free and open source software, it is using the LGPL-3.0 license.

See [LICENSE](LICENSE) for the full license text.

## FAQ

### MSEdgeRedirect isn't listed in the "How do you want to open this?" Menu?

Select the **Featured** Microsoft Edge. MSEdgeRedirect will still properly function.

### Will MSEdgeRedirect work with Edge uninstalled?

Not currently, see [#26](https://github.com/rcmaehl/MSEdgeRedirect/issues/26). The recommended work around is to install Edge Canary as it does not install system wide.

### It isn’t working for me?

Make sure the application is running in the system tray, then run `microsoft-edge:https://google.com` using the `Windows` + `R` keys. If that is not properly redirected, [file a bug report!](https://github.com/rcmaehl/MSEdgeRedirect/issues/new?assignees=&labels=&template=bug_report.md&title=)

### Will searches inside \<app name here\> still use Bing?

MSEdge Redirect only redirects links that attempt to open in MS Edge. It will not affect results generated within other applications.

### Can you change Bing results to Google Results?

Yes, as of 0.5.0.0, you can select One of 8 available Search Engines, or set your own!
  
### How do I uninstall?

**0.4.0.0 & Newer (also works with Winget and Chocolatey installation modes):** Use Programs and Features\
**BEFORE 0.4.0.0:** Simply delete MSEdgeRedirect.exe!

**Winget:**
```powershell
winget uninstall MSEdgeRedirect
```

**Chocolatey:**
```powershell
choco uninstall msedgeredirect
```
