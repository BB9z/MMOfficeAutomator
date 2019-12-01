
function Out-Info {
    <#
    .SYNOPSIS
    Log info text to the console.
    #>
    param (
        [string]$text
    )

    Write-Host -Object $text -ForegroundColor DarkGray
}
