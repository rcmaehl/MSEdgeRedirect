$ErrorActionPreference = 'Stop'

$packageArgs = @{
  packageName   = $env:ChocolateyPackageName
  softwareName  = 'MSEdgeRedirect'
  fileType      = 'exe'
  validExitCodes= @(@(0))
}

$uninstalled = $false

[array]$key = Get-UninstallRegistryKey @packageArgs

if ($key.Count -eq 1) {
  $key | ForEach-Object {
    $packageArgs['file'] = "$($_.UninstallString)"
    $fileStringSplit = $packageArgs['file'] -split '\s+(?=(?:[^"]|"[^"]*")*$)'

    if($fileStringSplit.Count -gt 1) {
      $packageArgs['file'] = $fileStringSplit[0]
      $packageArgs['silentArgs'] += " $($fileStringSplit[1..($fileStringSplit.Count-1)])"
    }

    Uninstall-ChocolateyPackage @packageArgs
  }
} elseif ($key.Count -eq 0) {
  Write-Warning "$packageName has already been uninstalled by other means."
} elseif ($key.Count -gt 1) {
  Write-Warning "$($key.Count) matches found!"
  Write-Warning "To prevent accidental data loss, no programs will be uninstalled."
  Write-Warning "Please alert the package maintainer that the following keys were matched:"
  $key | ForEach-Object { Write-Warning "- $($_.DisplayName)" }
}
