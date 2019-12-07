
using namespace System.Management
using namespace System.IO

. "$PSScriptRoot/Console.ps1"

class Workspace {

    [ValidateNotNullOrEmpty()][Automation.PathInfo]$Location
    [ValidateNotNullOrEmpty()][string[]]$OrganizationNames
    [ValidateNotNullOrEmpty()][string[]]$DocumentExtensions

    Workspace([Automation.PathInfo]$Location) {
        $this.Location = $Location
        $this.DocumentExtensions = @( ".doc", ".docx", ".xls", ".xlsx", ".pdf" )
    }

    [void]AddNumber() {
        $indexFromONameMap = @{ }
        $documents = $this.FilterDocumens(($this.Location | Get-ChildItem))
        $folders = $this.FilterFolders(($this.Location | Get-ChildItem))
        $total = 0
        foreach ($file in $documents + $folders) {
            if (!$file) { continue }
            $oName = $this.OrganizationNameFromFileName($file.BaseName)
            if (!$oName) {
                Write-Warning "  $($file.Name) 文件名不包括组织名，忽略"
                continue
            }
            $index = $indexFromONameMap[$oName]
            if (!$index) {
                $index = ($total++) + 1
                $indexFromONameMap[$oName] = $index
            }
            if ($file.Name -match "(?<number>\d+)\.*\s*(?<name>.+)") {
                if ([int]$Matches.number -eq $index) {
                    Write-Host "编号相同 $index"
                    [DConsole]::info("  $($file.Name) 已编号")
                }
                else {
                    Write-Host "编号不同"
                    $newName = "$($index.ToString('00')). $($Matches.name)"
                    [DConsole]::info(" 重新编号 $newName")
                    Rename-Item -Path $file -NewName $newName
                }
            }
            else {
                $newName = "$($index.ToString('00')). $($file.Name)"
                [DConsole]::info(" 编号 $newName")
                Rename-Item -Path $file -NewName $newName
            }
        }
    }

    [void]RemoveNumber() {
        $documents = $this.FilterDocumens(($this.Location | Get-ChildItem))
        $folders = $this.FilterFolders(($this.Location | Get-ChildItem))
        foreach ($file in $documents + $folders) {
            if (!$file) { continue }
            $oName = $this.OrganizationNameFromFileName($file.BaseName)
            if (!$oName) {
                Write-Warning "  $($file.Name) 文件名不包括组织名，忽略"
                continue
            }
            if ($file.Name -match "(?<number>\d+)\.*\s*(?<name>.+)") {
                $newName = $Matches.name
                [DConsole]::info(" 去除编号 $($newName)")
                Rename-Item -Path $file -NewName $newName
            }
            else {
                [DConsole]::info(" 忽略 $($file.Name)")
            }
        }
    }

    [void]PackageFileStruct() {
        $files = $this.FilterDocumens(($this.Location | Get-ChildItem))
        $folders = $this.FilterFolders(($this.Location | Get-ChildItem))
        $exsistFolderMap = @{ }
        foreach ($path in $folders) {
            $oName = $this.OrganizationNameFromFileName($path.Name)
            if ($oName) {
                $exsistFolderMap[$oName] = $path
            }
        }
        foreach ($file in $files) {
            $oName = $this.OrganizationNameFromFileName($file.BaseName)
            if ($oName) {
                $folder = $exsistFolderMap[$oName]
                if (!$folder) {
                    [DConsole]::info("  创建文件夹 $($oName)")
                    $folder = New-Item -Name $oName -ItemType "directory"
                    $exsistFolderMap[$oName] = $folder
                }
                [DConsole]::info("  移动 $($file.Name) => $($folder.Name)")
                Move-Item -Path $file -Destination $folder
            }
            else {
                Write-Warning "  $($file.Name) 文件名不包括组织名，无法自动归类"
            }
        }
    }

    [void]UnpackageFileStruct() {
        $folders = $this.FilterFolders(($this.Location | Get-ChildItem))
        foreach ($path in $folders) {
            $oName = $this.OrganizationNameFromFileName($path.Name)
            if (!$oName) {
                continue
            }
            foreach ($file in ($path | Get-ChildItem)) {
                [DConsole]::info("  提取 $($file.Name)")
                Move-Item $file -Destination $this.Location
            }
            [DConsole]::info("  删除文件夹 $($path.Name)")
            Remove-Item $path -Recurse -Force
        }
    }

    #region 工具

    [FileSystemInfo[]]FilterFolders([FileSystemInfo[]]$items) {
        $exceptFloder = @( "Scripts", ".vscode", ".git" )
        return $items | Where-Object {
            ($_.PSIsContainer) -and (!$exceptFloder.Contains($_.Name))
        }
    }

    [FileSystemInfo[]]FilterDocumens([FileSystemInfo[]]$items) {
        return $items | Where-Object {
            (!$_.PSIsContainer) -and ($this.DocumentExtensions.Contains($_.Extension))
        }
    }

    [string]OrganizationNameFromFileName([string]$FileName) {
        foreach ($item in $this.OrganizationNames) {
            if ($FileName.Contains($item)) {
                return $item
            }
        }
        return $null
    }
    #endregion
}
