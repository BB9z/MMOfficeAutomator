Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$config = Get-Content -Path .\config.json -Raw | ConvertFrom-Json
$DebugPreference = if ($config.debug) { "Continue" } else { "SilentlyContinue" }

. "$PSScriptRoot/Console.ps1"
. "$PSScriptRoot/Menu.ps1"
. "$PSScriptRoot/Workspace.ps1"

Write-Debug "launcher"
$Host.UI.RawUI.WindowTitle = "MMOffice"

$menu = [DMenu]::new("菜单", @(
        @{ID = "1"; Title = "任务1" },
        @{ID = "2"; Title = "任务2" },
        @{ID = "t"; Title = "测试" }
        @{ID = "q"; Title = "退出" }
    ), "请选择")

$shouldExit = $false
while (!$shouldExit) {
    Clear-Host
    switch ($menu.show()) {
        '1' {
            Write-Host "task 1"
            Pause
        }
        '2' {
            Write-Host "task 2"
            Pause
        }
        'q' {
            $shouldExit = $true
        }
    }
}

[DConsole]::Info("Bye!")
