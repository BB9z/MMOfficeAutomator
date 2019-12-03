Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$config = Get-Content -Path .\config.json -Raw | ConvertFrom-Json
$DebugPreference = if ($config.debug) { "Continue" } else { "SilentlyContinue" }

. "$PSScriptRoot/Console.ps1"
. "$PSScriptRoot/Workspace.ps1"

foreach ($item in $config.organization_list) {
    Out-Info $item
}

# $env:PSModulePaths
# New-Object System.IO.FileSystemWatchers

