[![Latest download count](https://img.shields.io/github/downloads/rcmaehl/MSEdgeRedirect/latest/total)](https://github.com/rcmaehl/MSEdgeRedirect/releases/latest/)
[![Chocolatey download count](https://img.shields.io/chocolatey/dt/msedgeredirect?label=Chocolatey+downloads)](https://chocolatey.org/packages/msedgeredirect)
[![Ko-fi](https://img.shields.io/badge/Support%20me%20on-Ko--fi-FF5E5B.svg?logo=ko-fi)](https://ko-fi.com/rcmaehl)
[![PayPal](https://img.shields.io/badge/Donate%20on-PayPal-00457C.svg?logo=paypal)](https://paypal.me/rhsky)
[![Join the Discord chat](https://img.shields.io/discord/728710400367001650?logo=discord)](https://discord.gg/uBnBcBx)

# MSEdgeRedirect
A Tool to Redirect News, Search, Widgets, Weather, and More to Your Default Browser

This tool filters and passes the command line arguments of Microsoft Edge processes into your default browser instead of hooking into the `microsoft-edge:` handler, this should provide resiliency against future changes. Additionally, an Image File Execution Options mode is available to operate similarly to the Old EdgeDeflector. Additional modes are planned for future versions.

No Default App walkthrough or other steps, just set and forget.

> :warning: PLEASE NOTE: MSEdgeRedirect is still BETA. Changes are to be expected, and performance to be improved.

## Recommended Alternatives

Looking for Alternatives? [Check Out This Chart](https://github.com/rcmaehl/MSEdgeRedirect/wiki/Alternative-Apps-Comparison-Chart)\
Not looking for Extra Features? Try [@AveYo](https://github.com/AveYo)'s [ChrEdgeFckOff](https://github.com/AveYo/fox/blob/main/ChrEdgeFkOff.cmd)\
Looking to Disable Web Search Entirely? Try [@krlvm](https://github.com/krlvm)'s [BeautySearch](https://github.com/krlvm/BeautySearch)

## Downloads

Download Stable (GitHub)|Download Testing (GitHub)
----|----
<a href="https://github.com/rcmaehl/MSEdgeRedirect/releases/latest/download/MSEdgeRedirect.exe"><img src="https://img.shields.io/github/v/release/rcmaehl/msedgeredirect?display_name=tag&style=for-the-badge" height="65px" /></a>|<a href="https://nightly.link/rcmaehl/MSEdgeRedirect/workflows/MSER/0.8.0.0-dev/mser.zip"><img src="https://img.shields.io/github/actions/workflow/status/rcmaehl/MSEdgeRedirect/MSER.yml?branch=0.8.0.0-dev&style=for-the-badge" height="65px" /></a>

### Package Managers

|<a href="https://community.chocolatey.org/packages/msedgeredirect/"><img src="https://user-images.githubusercontent.com/716581/159197666-761d9b5e-18f6-427c-bae7-2cc6bd348b9a.png" height="108px" /></a>|[![image](https://user-images.githubusercontent.com/716581/185218464-f84115df-fe0e-454c-9147-4da089273faf.png)](https://scoop.sh/#/apps?q=msedgeredirect&s=0&d=1&o=true)|[![image](https://user-images.githubusercontent.com/716581/159123573-58e5ccba-5c82-46ec-adcc-08b897284a6d.png)](https://github.com/microsoft/winget-pkgs/tree/master/manifests/r/rcmaehl/MSEdgeRedirect)|
|:----:|:----:|:----:|
|[![Chocolatey package](https://repology.org/badge/version-for-repo/chocolatey/msedgeredirect.svg)](https://repology.org/project/msedgeredirect/versions)|[![Scoop package](https://repology.org/badge/version-for-repo/scoop/msedgeredirect.svg)](https://repology.org/project/msedgeredirect/versions)|<!-- WINGET_PKG_START --><!-- WINGET_PKG_END -->|
|`choco install msedgeredirect`|`scoop bucket add extras`<br/>`scoop install msedgeredirect`|`winget install MSEdgeRedirect`|

### Compiling

1. Download and run "AutoIt Full Installation" from [official website](https://www.autoitscript.com/site/autoit/downloads). 
1. Get the source code either by [downloading zip](https://github.com/rcmaehl/MSEdgeRedirect/archive/main.zip) or do `git clone https://github.com/rcmaehl/MSEdgeRedirect`.
1. Right click on `MSEdgeRedirect.au3` in the MSEdgeRedirect directory and select "Compile Script (x64) (or x86 if you have 32 bit Windows install).
1. This will create MSEdgeRedirect.exe in the same directory.

### System Requirements
 |Minimum Requirements|Recommended
----|----|----
OS|Windows 8.1|Latest Windows 11 Build
CPU|32-bit Single Core|64-bit Dual Core or Higher
RAM (Memory)|40MB Free|100MB Free
Disk (Storage)|5MB Free|100MB Free

## Contributors

[![Contributors](https://contrib.rocks/image?repo=rcmaehl/MSEdgeRedirect)](https://github.com/rcmaehl/MSEdgeRedirect/graphs/contributors)

### Become a contributor

See [CONTRIBUTING.md](CONTRIBUTING.md) for rules of coding and pull requests.

## License

MSEdgeRedirect is free and open source software, it is using the LGPL-3.0 license.

See [LICENSE](LICENSE) for the full license text.

## FAQ

### MSEdgeRedirect isn't listed in the "How do you want to open this?" Menu?

Select the **Featured** Microsoft Edge. MSEdgeRedirect will still properly function.

### Will MSEdgeRedirect work with Edge uninstalled?

MSEdgeRedirect is compatible with [@AveYo](https://github.com/AveYo)'s [Edge Removal](https://github.com/AveYo/fox/blob/main/Edge_Removal.bat) as it retains a needed component. If Edge was removed using another method, reinstall Edge Stable, then run AveYo's tool. After AveYo's Edge Removal has been run, simply install MSEdgeRedirect and it will be detected automatically.

### It isn’t working for me?

Run `microsoft-edge:https://google.com` using the `Windows` + `R` keys. If that is not properly redirected, [file a bug report!](https://github.com/rcmaehl/MSEdgeRedirect/issues/new?assignees=&labels=&template=bug_report.md&title=)

### Will searches inside \<app name here\> still use Bing?

MSEdgeRedirect only redirects links that attempt to open in MS Edge. It will not affect results generated within other applications.

### Can you change Bing results to Google Results?

Yes, as of 0.5.0.0, you can select One of 8 available Search Engines, or set your own!
  
### How do I uninstall?

#### Normal Installs
Regular Install|Corrupted Install
----|----
Use Programs and Features|[Cleanup Tool](https://raw.githubusercontent.com/rcmaehl/MSEdgeRedirect/main/Assets/Cleanup%20Tool.ps1)

#### Package Managers
Chocolatey|Scoop|Winget
----|----|----
`choco uninstall msedgeredirect`|`scoop uninstall msedgeredirect`|`winget uninstall MSEdgeRedirect`
