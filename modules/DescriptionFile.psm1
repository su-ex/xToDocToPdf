[regex]$extractionPattern = '^(?<disabled>;?)(?<indent>\t*)(?<desc>.+):\s+((?<flags>[a-z0-9]*)>)?(?<path>[^:\*\?"<>\|]*)$'
[regex]$flagsExtractionPattern = '(?=.*a(?<alphabetical>r?))?(?=.*s(?<skiplang>))?(?=.*c(?<custombasepath>\d*))?(?=.*h(?<headingtier>[1-9]?))?.*'
$insertionPlaceholder = '${raw}'

[regex]$pdfFilenameInfoExtraction = '__(?<desc>.*)__(##(?<headingtier>[1-9])##)?'

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
        $alphabetical, $skipLang, $customBasePath, $headingTier = $f.Groups['alphabetical', 'skiplang', 'custombasepath', 'headingtier']

        # check if indentaion is always ascending with one step at max
        if ($indent.Length - $lastIndent -gt 1) {
            throw "Malformed description (indentation mistake)!"
        }
        $lastIndent = $indent.Length

        #interpret flags
        $flags = @{}
        If ($alphabetical.Success) {
            $flags["alphabetical"] = $alphabetical.Value -eq 'r'
        }
        If ($skipLang.Success) {
            $flags["skipLang"] = $Null
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
                $flags["headingTier"] = $indent.Length + 1
            } else {
                $flags["headingTier"] = $headingTier.Value
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

Function getDescribedFolder([string]$Path, [switch]$Recurse, [array]$Extensions) {
    $files = (
        Get-ChildItem $Path -Recurse -Name |
        ForEach-Object { makePathAbsolute $Path $_ } |
        Where-Object { $Extensions.Contains((Get-Item $_).Extension.ToLower()) } |
        Sort-Object
    )

    $piecesi = [System.Collections.ArrayList]@()
    foreach($file in $files) {
        $nameWithoutExtension = (Get-Item $file).BaseName
        
        # flag capture groups
        $f = $pdfFilenameInfoExtraction.Match($nameWithoutExtension)
        # $f | Format-List | Out-String | Write-Debug
        $desc, $headingTier = $f.Groups['desc', 'headingtier']

        $descValue = $nameWithoutExtension
        If ($desc.Success) {
            $descValue = $desc.Value
        }

        #interpret flags
        $flags = @{}
        If ($headingTier.Success) {
            $flags["headingTier"] = $headingTier.Value
        }

        # add each new successfully parsed entry to the list
        $piecesi.Add([PSCustomObject]@{
            raw = $Null
            desc = $descValue
            path = $file
            enabled = $true
            indent = -1
            flags = $flags
            asset = $Null
        }) | Out-Null
    }

    return $piecesi
}