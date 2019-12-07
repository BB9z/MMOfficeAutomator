
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
            $oName = $this.OrganizationNameFromFile($file)
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
            $oName = $this.OrganizationNameFromFile($file)
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
            $oName = $this.OrganizationNameFromFile($path)
            if ($oName) {
                $exsistFolderMap[$oName] = $path
            }
        }
        foreach ($file in $files) {
            $oName = $this.OrganizationNameFromFile($file)
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
            $oName = $this.OrganizationNameFromFile($path)
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

    [void]ReportStatistics() {
        $statisticMap = @{ }
        $files = $this.FilterDocumens(($this.Location | Get-ChildItem -Recurse))
        Write-Host "文件统计，文档共计: $($files.Count)"
        foreach ($file in $files) {
            $oName = $this.OrganizationNameFromFile($file)
            if (!$oName) {
                Write-Warning "  $($file.Name) 文件名不包括组织名，忽略"
                continue
            }
            $oInfo = $statisticMap[$oName]
            if (!$oInfo) {
                $oInfo = @{
                    HasDoc = $false;
                    HasXls = $false;
                    HasOther = $false;
                    OtherCount = 0;
                }
                $statisticMap[$oName] = $oInfo
            }
            switch ($file.Extension) {
                { $_ -eq ".doc" -or $_ -eq ".docx" } {
                    $oInfo.HasDoc = $true
                }
                { $_ -eq ".xls" -or $_ -eq ".xlsx" } {
                    $oInfo.HasXls = $true
                }
                Default {
                    $oInfo.HasOther = $true
                    $oInfo.OtherCount += 1
                }
            }
        }

        $organizations = $this.OrganizationNames
        $noAnyOrgs = [System.Collections.ArrayList]$organizations.Clone()
        $noXlsOrgs = [System.Collections.ArrayList]$organizations.Clone()
        $hasXlsOrgs = @()
        foreach ($oName in $statisticMap.Keys) {
            $oInfo = $statisticMap[$oName]
            $noAnyOrgs.Remove($oName)
            if ($oInfo.HasXls) {
                $noXlsOrgs.Remove($oName)
                $hasXlsOrgs += $oName
            }
        }

        if ($noAnyOrgs -and $noAnyOrgs.Count -gt 0) {
            Write-Host "`n无任何文件的组织有 $($noAnyOrgs.Count)家" -ForegroundColor "Red"
            foreach ($oName in $noAnyOrgs) {
                Write-Host "  $oName" -ForegroundColor DarkRed
            }
        }

        if ($noXlsOrgs -and $noXlsOrgs.Count -gt 0) {
            Write-Host "`n$($hasXlsOrgs.Count) 家有报表，无报表 $($noXlsOrgs.Count) 家：" -ForegroundColor "Yellow"
            foreach ($oName in $noXlsOrgs) {
                Write-Host "  $oName" -ForegroundColor DarkRed
            }
        }
        else {
            Write-Host "`n报表全部已报" -ForegroundColor "Green"
        }

        Write-Host "`n$($organizations.Count) 家组织详情："
        foreach ($oName in $organizations) {
            Write-Host "  $($oName)`t" -NoNewline
            $oInfo = $statisticMap[$oName]
            if (!$oInfo) {
                Write-Host "无任何文件" -ForegroundColor "Red"
                continue
            }
            if ($oInfo.HasXls) {
                Write-Host "  有报表" -NoNewline -ForegroundColor "Green"
            } else {
                Write-Host "  无报表" -NoNewline -ForegroundColor "Yellow"
            }
            if ($oInfo.HasDoc) {
                Write-Host "  有 Doc" -NoNewline -ForegroundColor "Blue"
            } else {
                Write-Host "  无 Doc" -NoNewline -ForegroundColor "Yellow"
            }
            if ($oInfo.HasOther) {
                Write-Host "  其他文件 $($oInfo.OtherCount) 个"
            }
            else {
                Write-Host ""
            }
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
            (!$_.PSIsContainer) -and ($this.DocumentExtensions.Contains($_.Extension.ToLower()))
        }
    }

    [string]OrganizationNameFromFile([FileSystemInfo]$File) {
        $name = $File.BaseName
        foreach ($item in $this.OrganizationNames) {
            if ($name.Contains($item)) {
                return $item
            }
        }
        return $null
    }
    #endregion
}
