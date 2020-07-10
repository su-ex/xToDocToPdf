Param (
    [String] ${working-directory},
    
    [String] ${source-word-file} = ".\target.docx"
)

Import-Module "$PSScriptRoot\modules\WordPdfExportHelper.psm1" -Force
Import-Module "$PSScriptRoot\modules\HelperFunctions.psm1" -Force

# enforce working-directory to be passed as parameter
if (-not $PSBoundParameters.ContainsKey('working-directory')) {
    exitError "Ein Arbeitsverzeichnis muss als Parameter übergeben werden!"
}

# make sure working directory exists and make path always absolute
try {
    $Script:workingDirectory = Resolve-Path ${working-directory} -ErrorAction Stop
} catch {
    exitError "Das Arbeitsverzeichnis existiert nicht: $_"
}

# base source file upon working directory
$sourceWordFile = makePathAbsolute $workingDirectory ${source-word-file}
if (-not (Test-Path $sourceWordFile)) {
    exitError "Das Quelldokument $sourceWordFile existiert nicht!"
}

$wpeh = WordPdfExportHelper $sourceWordFile

try {
    $replacements = $wpeh.getPdfReplacementPages()
    $totalPageCount = $wpeh.getTotalPageCount()

    $list = [System.Collections.ArrayList]@()
    for ($i = 0; $i -lt $totalPageCount; $i++) {
        $list.Add($true) | Out-Null
    }

    foreach ($replacement in $replacements) {
        $list[$replacement.docPageNumber-1] = $false
    }

    $ranges = [System.Collections.ArrayList]@()
    $i = 1
    $range = @(0, 0)
    $didadd = $false
    while ($i -le $list.Count) {
        if ($list[$i-1] -eq $true) {
            if ($range[0] -eq 0) {
                $range[0] = $i
                $range[1] = $i
                $didadd = $false
            } elseif ($i -eq $range[1]+1) {
                $range[1] = $i
                $didadd = $false
            }
        } elseif (-not $didadd) {
            $ranges.Add($range) | Out-Null
            $range = @(0, 0)
            $didadd = $true
        }
        $range[1] = $i
        $i += 1
    }
    if (-not $didadd) {
        $ranges.Add($range) | Out-Null
    }

    $s = (($ranges | ForEach-Object {($_[0] -eq $_[1]) ? "$($_[0])" : ($_ -join "-")}) -join ",")
    Write-Host $s
    Set-Clipboard $s
    infoBox "Ist in Zwischenablage!"
} catch {
    $_ | Out-String | Write-Error
    exitError "Leider fehlgeschlagen: $($_.Exception.Message)"
} finally {
    $wpeh.destroy()
}

exit 0
