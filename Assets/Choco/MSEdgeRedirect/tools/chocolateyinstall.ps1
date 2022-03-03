$toolsPath = Split-Path -parent $MyInvocation.MyCommand.Definition

$packageArgs = @{
  packageName    = $env:ChocolateyPackageName
  fileType       = 'exe'
  url            = '{URL32}'
  checksum       = '{SHA256CHECKSUM32}'
  checksumType   = 'sha256'
  url64Bit       = '{URL64}'
  checksum64     = '{SHA256CHECKSUM64}'
  checksumType64 = 'sha256'
  softwareName   = 'MSEdgeRedirect'
  silentArgs     = '/wingetinstall'
  validExitCodes = @(0)
}

Install-ChocolateyPackage @packageArgs
