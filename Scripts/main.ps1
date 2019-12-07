Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$config = Get-Content -Path .\config.json -Raw | ConvertFrom-Json
$DebugPreference = if ($config.debug) { "Continue" } else { "SilentlyContinue" }

. "$PSScriptRoot/Console.ps1"
. "$PSScriptRoot/Menu.ps1"
. "$PSScriptRoot/Workspace.ps1"

Write-Debug "launcher"
$Host.UI.RawUI.WindowTitle = "MMOffice"

$work = [Workspace]::new((Get-Location))
$work.OrganizationNames = $config.organization_list

$menu = [DMenu]::new("菜单", @(
        @{ID = "1"; Title = "编号添加" },
        @{ID = "2"; Title = "编号去除" },
        @{ID = "3"; Title = "文件归类到文件夹" }
        @{ID = "4"; Title = "把文件夹中的文件提取出来" }
        @{ID = "q"; Title = "退出" }
    ), "请选择")

$shouldExit = $false
while (!$shouldExit) {
    Clear-Host
    switch ($menu.show()) {
        '1' {
            $work.AddNumber()
            Pause
        }
        '2' {
            $work.RemoveNumber()
            Pause
        }
        '3' {
            $work.PackageFileStruct()
            Pause
        }
        '4' {
            $work.UnpackageFileStruct()
            Pause
        }
        'q' {
            $shouldExit = $true
        }
    }
}

[DConsole]::Info("Bye!")
