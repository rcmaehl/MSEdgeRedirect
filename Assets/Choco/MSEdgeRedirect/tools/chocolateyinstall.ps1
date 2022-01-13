# IMPORTANT: Before releasing this package, copy/paste the next 2 lines into PowerShell to remove all comments from this file:
#   $f='c:\path\to\thisFile.ps1'
#   gc $f | ? {$_ -notmatch "^\s*#"} | % {$_ -replace '(^.*?)\s*?[^``]#.*','$1'} | Out-File $f+".~" -en utf8; mv -fo $f+".~" $f

# 1. See the _TODO.md that is generated top level and read through that
# 2. Follow the documentation below to learn how to create a package for the package type you are creating.
# 3. In Chocolatey scripts, ALWAYS use absolute paths - $toolsDir gets you to the package's tools directory.
$ErrorActionPreference = 'Stop'; # stop on all errors
$toolsDir   = "$(Split-Path -parent $MyInvocation.MyCommand.Definition)"
# Internal packages (organizations) or software that has redistribution rights (community repo)
# - Use `Install-ChocolateyInstallPackage` instead of `Install-ChocolateyPackage`
#   and put the binaries directly into the tools folder (we call it embedding)
$fileLocation = Join-Path $toolsDir 'MSEdgeRedirect.exe'
# If embedding binaries increase total nupkg size to over 1GB, use share location or download from urls
#$fileLocation = '\\SHARE_LOCATION\to\INSTALLER_FILE'
# Community Repo: Use official urls for non-redist binaries or redist where total package size is over 200MB
# Internal/Organization: Download from internal location (internet sources are unreliable)
$url        = '' # download url, HTTPS preferred
$url64      = '' # 64bit URL here (HTTPS preferred) or remove - if installer contains both (very rare), use $url

$packageArgs = @{
  packageName   = $env:ChocolateyPackageName
  unzipLocation = $toolsDir
  fileType      = 'EXE' #only one of these: exe, msi, msu
  #url           = $url
  #url64bit      = $url64
  #file         = $fileLocation

  softwareName  = 'MSEdgeRedirect*' #part or all of the Display Name as you see it in Programs and Features. It should be enough to be unique

  # Checksums are now required as of 0.10.0.
  # To determine checksums, you can get that from the original site if provided.
  # You can also use checksum.exe (choco install checksum) and use it
  # e.g. checksum -t sha256 -f path\to\file
  checksum      = ''
  checksumType  = 'sha256' #default is md5, can also be sha1, sha256 or sha512
  checksum64    = ''
  checksumType64= 'sha256' #default is checksumType

  # MSI
  # silentArgs    = "/qn /norestart /l*v `"$($env:TEMP)\$($packageName).$($env:chocolateyPackageVersion).MsiInstall.log`"" # ALLUSERS=1 DISABLEDESKTOPSHORTCUT=1 ADDDESKTOPICON=0 ADDSTARTMENU=0
  validExitCodes= @(0)
}

#Install-ChocolateyPackage @packageArgs # https://docs.chocolatey.org/en-us/create/functions/install-chocolateypackage
#Install-ChocolateyZipPackage @packageArgs # https://docs.chocolatey.org/en-us/create/functions/install-chocolateyzippackage
## If you are making your own internal packages (organizations), you can embed the installer or
## put on internal file share and use the following instead (you'll need to add $file to the above)
Install-ChocolateyInstallPackage @packageArgs # https://docs.chocolatey.org/en-us/create/functions/install-chocolateyinstallpackage

## Main helper functions - these have error handling tucked into them already
## see https://docs.chocolatey.org/en-us/create/functions

## Install an application, will assert administrative rights
## - https://docs.chocolatey.org/en-us/create/functions/install-chocolateypackage
## - https://docs.chocolatey.org/en-us/create/functions/install-chocolateyinstallpackage
## add additional optional arguments as necessary
##Install-ChocolateyPackage $packageName $fileType $silentArgs $url [$url64 -validExitCodes $validExitCodes -checksum $checksum -checksumType $checksumType -checksum64 $checksum64 -checksumType64 $checksumType64]

## Installer, will assert administrative rights - used by Install-ChocolateyPackage
## use this for embedding installers in the package when not going to community feed or when you have distribution rights
## - https://docs.chocolatey.org/en-us/create/functions/install-chocolateyinstallpackage
#Install-ChocolateyInstallPackage $packageName $fileType $silentArgs '_FULLFILEPATH_' -validExitCodes $validExitCodes

## Runs processes asserting UAC, will assert administrative rights - used by Install-ChocolateyInstallPackage
## - https://docs.chocolatey.org/en-us/create/functions/start-chocolateyprocessasadmin
#Start-ChocolateyProcessAsAdmin 'STATEMENTS_TO_RUN' 'Optional_Application_If_Not_PowerShell' -validExitCodes $validExitCodes

## To avoid quoting issues, you can also assemble your -Statements in another variable and pass it in
#$appPath = "$env:ProgramFiles\appname"
##Will resolve to C:\Program Files\appname
#$statementsToRun = "/C `"$appPath\bin\installservice.bat`""
#Start-ChocolateyProcessAsAdmin $statementsToRun cmd -validExitCodes $validExitCodes

## add specific folders to the path - any executables found in the chocolatey package
## folder will already be on the path. This is used in addition to that or for cases
## when a native installer doesn't add things to the path.
## - https://docs.chocolatey.org/en-us/create/functions/install-chocolateypath
#Install-ChocolateyPath 'LOCATION_TO_ADD_TO_PATH' 'User_OR_Machine' # Machine will assert administrative rights

## Add specific files as shortcuts to the desktop
## - https://docs.chocolatey.org/en-us/create/functions/install-chocolateyshortcut
#$target = Join-Path $toolsDir "$($packageName).exe"
# Install-ChocolateyShortcut -shortcutFilePath "<path>" -targetPath "<path>" [-workDirectory "C:\" -arguments "C:\test.txt" -iconLocation "C:\test.ico" -description "This is the description"]

## Outputs the bitness of the OS (either "32" or "64")
## - https://docs.chocolatey.org/en-us/create/functions/get-osarchitecturewidth
#$osBitness = Get-ProcessorBits

## Set up a file association
## - https://docs.chocolatey.org/en-us/create/functions/install-chocolateyfileassociation
#Install-ChocolateyFileAssociation

## Adding a shim when not automatically found - Cocolatey automatically shims exe files found in package directory.
## - https://docs.chocolatey.org/en-us/create/functions/install-binfile
## - https://docs.chocolatey.org/en-us/create/create-packages#how-do-i-exclude-executables-from-getting-shims
#Install-BinFile