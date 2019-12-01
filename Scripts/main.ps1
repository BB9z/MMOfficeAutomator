$config = Get-Content -Path .\config.json -Raw | ConvertFrom-Json

if ($config.debug) {
    $DebugPreference = "Continue"
}
else {
    $DebugPreference = "SilentlyContinue"
}

. "$PSScriptRoot/Log.ps1"

foreach ($item in $config.organization_list) {
    Out-Info $item
}

# $env:PSModulePaths
# New-Object System.IO.FileSystemWatchers

