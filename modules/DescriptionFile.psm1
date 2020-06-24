[Regex]$extractionPattern = '^(?<disabled>;?)(?<indent>\t*)(?<desc>.+):\s+(?<path>[^\\/:\*\?"<>\|]+)$'
$insertionPlaceholder = '${disabled}${indent}${desc}: ${path}'

Import-Module "$PSScriptRoot\HelperFunctions.psm1" -Force

Function getDescription($path) {
    $pieces = [System.Collections.ArrayList]@()

    $lastIndent = 0;
    foreach($line in [System.IO.File]::ReadLines($path)) {
        $m = $extractionPattern.match($line)
        $m | Format-List | Out-String | Write-Debug

        if (-not $m) {
            throw "Malformed description (pattern not recognized)!"
        }

        $disabled, $indent, $desc, $path = $m.Groups['disabled', 'indent', 'desc', 'path']

        if ($indent.Length - $lastIndent -gt 1) {
            throw "Malformed description (indentation mistake)!"
        }
        $lastIndent = $indent.Length

        $pieces.Add([PSCustomObject]@{
            desc = $desc.Value
            path = $path.Value
            enabled = $disabled.Length -eq 0
            indent = $indent.Length
            asset = $Null
        }) | Out-Null
    }

    return $pieces
}

function setDescription($path, $description) {
    $description | ForEach-Object {
        replaceTokens $insertionPlaceholder @{
            disabled = IIf $_.enabled "" ";"
            indent = "`t" * $_.indent
            desc = $_.desc
            path = $_.path
        }
    } | Out-File -FilePath $path
}