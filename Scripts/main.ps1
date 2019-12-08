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

function StartFileMonitor () {
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = Get-Location
    $watcher.IncludeSubdirectories = $false
    $watcher.EnableRaisingEvents = $true

    $EventType = @{
        Changed = 'Changed';
        Created = 'Created';
        Renamed = 'Renamed';
        Deleted = 'Deleted';
    }

    try {
        $action = {
            $details = $event.SourceEventArgs
            $changeType = $details.ChangeType

            if ($details.Name -match "^[~_.].*") {
                Write-Host "忽略非正常文件 $($details.Name)"
                return
            }

            $changePath = $details.FullPath
            [DConsole]::info("$($event.TimeGenerated): $changeType $changePath")

            if ($changeType -eq 'Created') {
                $path = Get-Item "$changePath"

                $oName = $work.OrganizationNameFromFile($path)
                if (!$oName) {
                    Write-Host " $newName 不包含组织名"
                    return
                }
                Write-Host "发现属于 $oName 的新文件"

                $newName = "$($oName)$($path.Extension)"
                if ($path.Name -eq $newName) {
                    Write-Host "正常命名，无需重命名"
                }
                elseif (Test-Path -Path $newName) {
                    Write-Warning " $newName 已存在，跳过重命名"
                }
                else {
                    [DConsole]::info("重命名 $($path.Name) => $newName")
                    Rename-Item -Path "$path" -NewName "$newName"
                }
            }
        }

        $handlers = . {
            Register-ObjectEvent -InputObject $watcher -EventName $EventType.Created -Action $action -SourceIdentifier 'FSCreate'
            # Register-ObjectEvent -InputObject $watcher -EventName $EventType.Renamed -Action $action -SourceIdentifier 'FSRename'
            Register-ObjectEvent -InputObject $watcher -EventName $EventType.Deleted -Action $action -SourceIdentifier 'FSDelete'
        }

        $continue = $true
        do {
            if ([System.Console]::KeyAvailable) {
                $x = [System.Console]::ReadKey()
                switch ($x.key) {
                    'q' { $continue = $false }
                }
            }
            Wait-Event -Timeout 1
        } while ($continue)
    }
    catch {
        $_ | Write-Error
    }
    finally {
        Unregister-Event -SourceIdentifier 'FSCreate'
        # Unregister-Event -SourceIdentifier 'FSRename'
        Unregister-Event -SourceIdentifier 'FSDelete'
        $handlers | Remove-Job
        $watcher.EnableRaisingEvents = $false
        $watcher.Dispose()
    }
}

$menu = [DMenu]::new("菜单", @(
        @{ID = "1"; Title = "编号添加" },
        @{ID = "2"; Title = "编号去除" },
        @{ID = "3"; Title = "文件归类到文件夹" }
        @{ID = "4"; Title = "把文件夹中的文件提取出来" }
        @{ID = "5"; Title = "统计" }
        @{ID = "6"; Title = "监听文件夹，自动重命名" }
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
        '5' {
            $work.ReportStatistics()
            Pause
        }
        '6' {
            Clear-Host
            Write-Host "开始监听 $(Get-Location) 目录"
            Write-Host "  按 Q 以停止监听" -ForegroundColor "Gray"
            StartFileMonitor
        }
        'q' {
            $shouldExit = $true
        }
    }
}

[DConsole]::Info("Bye!")
