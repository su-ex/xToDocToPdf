[Regex]$extractionPattern = '^(?<disabled>;?)(?<indent>\t*)(?<desc>.+):(\s+((?=.*a(?<alphabetical>r?))?(?=.*c(?<custombasepath>\d+))?(?=.*h(?<headingtier>\d))?.*>)?(?<path>[^\\/:\*\?"<>\|]+))?$'
$insertionPlaceholder = '${disabled}${indent}${desc}: ${path}'

Import-Module "$PSScriptRoot\HelperFunctions.psm1" -Force

Function getDescription($path) {
    $pieces = [System.Collections.ArrayList]@()

    $lastIndent = 0;
    foreach($line in [System.IO.File]::ReadLines($path)) {
        $m = $extractionPattern.match($line)
        $m | Format-List | Out-String | Write-Debug

        if (-not $m.Success) {
            throw "Malformed description (pattern not recognized)!"
        }

        $disabled, $indent, $desc, $path = $m.Groups['disabled', 'indent', 'desc', 'path'] # basic capture groups
        $alphabetical, $customBasePath, $headingTier = $m.Groups['alphabetical', 'custombasepath', 'headingtier'] # flag capture groups

        if ($indent.Length - $lastIndent -gt 1) {
            throw "Malformed description (indentation mistake)!"
        }
        $lastIndent = $indent.Length

        $flags = @{}
        If ($alphabetical.Success) {
            $flags["alphabetical"]   = $alphabetical.Value -eq 'r'
        }
        If ($customBasePath.Success) {
            $flags["customBasePath"] = [int]$customBasePath.Value
            If ($flags["customBasePath"] -gt 0) {
                $flags["customBasePath"] -= 1
            }
        }
        If ($headingTier.Success) {
            $flags["headingTier"]    = $headingTier.Value
        }

        $pieces.Add([PSCustomObject]@{
            desc = $desc.Value
            path = $path.Value
            enabled = $disabled.Length -eq 0
            indent = $indent.Length
            flags = $flags
            asset = $Null
        }) | Out-Null
    }

    return $pieces
}

Function setDescription($path, $description) {
    $description | ForEach-Object {
        replaceTokens $insertionPlaceholder @{
            disabled = IIf $_.enabled "" ";"
            indent = "`t" * $_.indent
            desc = $_.desc
            path = $_.path
        }
    } | Out-File -FilePath $path
}

Function getDescribedFolder() {}