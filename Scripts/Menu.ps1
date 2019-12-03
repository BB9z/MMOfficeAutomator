
. "$PSScriptRoot/Console.ps1"

class DMenu {

    [ValidateNotNullOrEmpty()][string]$Title
    [ValidateNotNullOrEmpty()][string]$Promote
    [Object]$Items

    DMenu(
        [string]$Title,
        [Object]$MenuItems,
        [string]$Promote
    ) {
        $this.Title = $Title
        $this.Promote = if ($Promote) { $Promote } else { "请输入你的选择" }
        $this.Items = $MenuItems
    }

    [void]AddItem([string]$ID, [string]$Title) {
        $this.Items += @{ ID = $ID; Title = $Title }
    }

    [string]Show() {
        Write-Host $this.Title

        $vaildInput = @()
        foreach ($item in $this.Items) {
            Write-Host " $($item.ID)" -NoNewline -ForegroundColor Yellow
            Write-Host ". " -NoNewline -ForegroundColor Gray
            Write-Host "$($item.Title)" -ForegroundColor DarkGreen
            $vaildInput += "$($item.ID)"
        }

        $in = Read-Host -Prompt "$($this.Promote)"
        if ($vaildInput.Length -eq 0) {
            return $in
        }

        $color = "Red"
        while (!$vaildInput.Contains($in)) {
            [DConsole]::ResetLastLine()
            Write-Host "`r$($this.Promote)" -NoNewline
            $color = if ($color -eq "Red") { "DarkRed" } else { "Red" }
            Write-Host "（请输入序号）" -NoNewline -ForegroundColor $color
            Write-Host ": " -NoNewline
            $in = Read-Host
        }
        return $in
    }
}
