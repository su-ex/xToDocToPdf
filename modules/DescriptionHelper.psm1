[regex]$extractionPattern = '^(?<disabled>;?)(?<indent>\t*)(?<desc>.+):\s+(?<rawflags>((?<flags>[a-z0-9]*)>))?(?<path>[^:\*\?"<>\|]*)$'
[regex]$flagsExtractionPattern = '^(?=.*a(?<alphabetical>r?))?(?=.*s(?<skiplang>))?(?=.*c(?<custombasepath>\d*))?(?=.*h(?<headingtier>[1-9r]?))?.*$'
$insertionPlaceholder = '${disabled}${indent}${desc}: ${rawflags}${path}'

[regex]$pdfFilenameInfoExtraction = '^(?=.*__(?<desc>.*)__)?(?=.*##(?<headingtier>[1-9r]?)##)?.*$'
$pdfFilenameInfoReplace = @(@('##.*##', ''))

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
        $disabled, $indent, $desc, $rawFlags, $path = $m.Groups['disabled', 'indent', 'desc', 'rawflags', 'path']
        
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
            $flags["descOnly"] = $true
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
        if ($path.Value -eq "") {
            $flags["descOnly"] = $true
        }

        Write-Debug "flags:"
        $flags | Format-List | Out-String | Write-Debug 

        # add each new successfully parsed entry to the list
        $pieces.Add([PSCustomObject]@{
            desc = $desc.Value
            path = $path.Value
            enabled = $disabled.Length -eq 0
            indent = $indent.Length
            rawFlags = $rawFlags.Value
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
            rawflags = $_.rawFlags
            path = $_.path
        }
    } | Out-File -FilePath $path
}

Function getDescribedFolder([string]$Path, [switch]$Recurse, [array]$Extensions, [int]$Indent) {
    $files = (
        Get-ChildItem $Path -Name |
        Sort-Object |
        ForEach-Object { makePathAbsolute $Path $_ }
    )
    
    $pieces = [System.Collections.ArrayList]@()
    foreach($file in $files) {
        #add flags
        $flags = @{}

        $item = Get-Item $file
        $nameWithoutExtension = $item.BaseName
        while ($item.Extension -eq ".lnk") {
            $file = Get-ShortcutTargetPath $item
            $item = Get-Item $file
            $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($nameWithoutExtension)
        }

        $addSubPieces = $false
        if ($Extensions.Contains($item.Extension.ToLower())) {
            # go on
        } elseif ($Recurse.IsPresent -and $item -is [System.IO.DirectoryInfo]) {
            $addSubPieces = $true
            $flags["descOnly"] = $true
        } else {
            # skip this
            continue
        }
        
        # capture info from filename
        $f = $pdfFilenameInfoExtraction.Match($nameWithoutExtension)
        $desc, $headingTier = $f.Groups['desc', 'headingtier']

        # use desc text captured from filename, otherwise use filename itself 
        $descValue = $nameWithoutExtension
        If ($desc.Success) {
            $descValue = $desc.Value
        } else {
            $descValue = replaceEachInString $descValue $pdfFilenameInfoReplace
        }

        If ($headingTier.Success) {
            $flags["headingTier"] = $headingTier.Value
        }

        # add each new successfully parsed entry to the list
        $pieces.Add([PSCustomObject]@{
            desc = $descValue
            path = $file
            enabled = $true
            indent = $Indent
            flags = $flags
            asset = $Null
        }) | Out-Null

        if ($addSubPieces) {
            foreach ($sp in (getDescribedFolder -Path $file -Recurse:($Recurse) -Extensions $Extensions -Indent ($Indent + 1))) {
                $pieces.Add($sp) | Out-Null
            }
        }
    }

    return $pieces
}