[Regex]$extractionPattern = '^(?<disabled>;?)(?<indent>\t*)(?<desc>.+):\s+(?<rawflags>((?=.*a(?<alphabetical>r?))?(?=.*c(?<custombasepath>\d+))?(?=.*h(?<headingtier>\d*))?.*>)?)(?<path>[^:\*\?"<>\|]*)$'
$insertionPlaceholder = '${disabled}${indent}${desc}: ${rawflags}${path}'

Import-Module "$PSScriptRoot\HelperFunctions.psm1" -Force

Function getDescription($path) {
    $pieces = [System.Collections.ArrayList]@()

    $i = 1
    $lastIndent = 0
    foreach($line in [System.IO.File]::ReadAllLines($path)) {
        $m = $extractionPattern.match($line)
        Write-Debug "line: $line"
        $m | Format-List | Out-String | Write-Debug

        # pattern not recognized
        if (-not $m.Success) {
            throw "Malformed description (pattern not recognized at line $i):`n$line"
        }

        $disabled, $indent, $desc, $path = $m.Groups['disabled', 'indent', 'desc', 'path'] # basic capture groups
        $rawFlags, $alphabetical, $customBasePath, $headingTier = $m.Groups['rawflags', 'alphabetical', 'custombasepath', 'headingtier'] # flag capture groups

        # check if indentaion is always ascending with one step at max
        if ($indent.Length - $lastIndent -gt 1) {
            throw "Malformed description (indentation mistake)!"
        }
        $lastIndent = $indent.Length

        #interpret flags
        $flags = @{"raw" = $rawFlags}
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
            if ($headingTier.Value -eq "") {
                $flags["headingTier"]    = $indent.Length
            } else {
                $flags["headingTier"]    = $headingTier.Value
            }
        }

        Write-Debug "flags:"
        $flags | Format-List | Out-String | Write-Debug 

        # add each new successfully parsed entry to the list
        $pieces.Add([PSCustomObject]@{
            desc = $desc.Value
            path = $path.Value
            enabled = $disabled.Length -eq 0
            indent = $indent.Length
            flags = $flags
            asset = $Null
        }) | Out-Null

        $i += 1
    }

    return $pieces
}

Function setDescription($path, $description) {
    $description | ForEach-Object {
        replaceTokens $insertionPlaceholder @{
            disabled = IIf $_.enabled "" ";"
            indent = "`t" * $_.indent
            desc = $_.desc
            rawflags = $_.flags.raw
            path = $_.path
        }
    } | Out-File -FilePath $path
}

Function getDescribedFolder() {}