$packageArgs = @{
  packageName    = $env:ChocolateyPackageName
  fileType       = 'exe'
  url64Bit       = '{URL64}'
  checksum64     = '{SHA256CHECKSUM64}'
  checksumType64 = 'sha256'
  softwareName   = 'MSEdgeRedirect'
  silentArgs     = '/wingetinstall'
  validExitCodes = @(0)
}

Install-ChocolateyPackage @packageArgs
