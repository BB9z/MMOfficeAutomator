
class DOut {
    <#
    .SYNOPSIS
    Inspect object.
    #>
    static Inspect($object) {
        $object.GetType().FullName | Write-Host
        Write-Host ($object | Format-List -Force | Out-String)
    }

    <#
    .SYNOPSIS
    Inspect object interface.
    #>
    static Inspect2($object) {
        $object | Get-Member | Out-Host -Paging
    }
}

class DConsole {
    static [System.ConsoleKeyInfo]ReadKey([bool]$noEcho = $true) {
        $key = [System.Console]::ReadKey()
        if ($noEcho) {
            $cTop = [System.Console]::CursorTop
            [System.Console]::SetCursorPosition(0, $cTop)
        }
        return $key
    }

    # Read-Host -Prompt "Enter text"

    static [void]info($input) {
        $input | Write-Host -ForegroundColor DarkGray
    }

    static [void]ResetLastLine() {
        $cTop = [System.Console]::CursorTop
        [System.Console]::SetCursorPosition(0, $cTop - 1)
        $cWidth = (Get-Host).UI.RawUI.BufferSize.Width
        [System.Console]::Write("{0,-$cWidth}" -f " ")
    }
}

function Write-Info {
    <#
    .SYNOPSIS
    Log info text to the console.
    #>

    process {
        $args | Write-Host -ForegroundColor DarkGray
        $input | Write-Host -ForegroundColor DarkGray
    }
}
