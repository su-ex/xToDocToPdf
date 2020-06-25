[regex]$extractionPattern = '^(?<disabled>;?)(?<indent>\t*)(?<desc>.+):\s+((?<flags>[a-z0-9]*)>)?(?<path>[^:\*\?"<>\|]*)$'
[regex]$flagsExtractionPattern = '(?=.*a(?<alphabetical>r?))?(?=.*c(?<custombasepath>\d*))?(?=.*h(?<headingtier>[1-9]?))?.*'
$insertionPlaceholder = '${raw}'

Import-Module "$PSScriptRoot\HelperFunctions.psm1" -Force

Function getDescription($path) {
    $pieces = [System.Collections.ArrayList]@()

    $i = 1
    $lastIndent = 0
    foreach($line in [System.IO.File]::ReadAllLines($path)) {
        $m = $extractionPattern.match($line)
        Write-Debug "line: $line"
        # $m | Format-List | Out-String | Write-Debug

        # pattern not recognized
        if (-not $m.Success) {
            throw "Malformed description (pattern not recognized at line $i):`n$line"
        }
        
        # basic capture groups
        $disabled, $indent, $desc, $path = $m.Groups['disabled', 'indent', 'desc', 'path']
        
        # flag capture groups
        $f = $flagsExtractionPattern.Match($m.Groups['flags'].Value)
        # $f | Format-List | Out-String | Write-Debug
        $alphabetical, $customBasePath, $headingTier = $f.Groups['alphabetical', 'custombasepath', 'headingtier']

        # check if indentaion is always ascending with one step at max
        if ($indent.Length - $lastIndent -gt 1) {
            throw "Malformed description (indentation mistake)!"
        }
        $lastIndent = $indent.Length

        #interpret flags
        $flags = @{}
        If ($alphabetical.Success) {
            $flags["alphabetical"]   = $alphabetical.Value -eq 'r'
        }
        If ($customBasePath.Success) {
            if ($customBasePath.Value -eq "") {
                $flags["customBasePath"] = 0
            } else {
                $flags["customBasePath"] = ([int]$customBasePath.Value) - 1
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
            raw = $line
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
            raw = $_.raw
        }
    } | Out-File -FilePath $path
}

Function getDescribedFolder() {}