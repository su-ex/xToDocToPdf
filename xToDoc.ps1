﻿Param (
    [String] ${working-directory} = ".",
    
    [String] ${target-file} = ".\target.docx",

    [Switch] ${no-selected-description-file},
    [String] ${selected-description-file},

    [String] ${template-description-file},

    [Switch] ${skip-tree-selection},

    [String] $lang,

    [String[][]] $variables,
    
    [Switch] ${get-variables-from-excel},
    [String] ${excel-variables-workbook-file},
    [String] ${excel-variables-worksheet-name},
    [String] ${excel-variables-table-name},
    
    [Switch] ${get-translations-from-excel},
    [String] ${excel-translations-workbook-file},
    [String] ${excel-translations-worksheet-name},
    [String] ${excel-translations-table-name},

    [String] ${custom-template-pdf-page},

    [String[]] ${custom-base-path} = @(".")
)

Import-Module "$PSScriptRoot\modules\WordAbstraction.psm1" -Force
Import-Module "$PSScriptRoot\modules\DescriptionHelper.psm1" -Force
Import-Module "$PSScriptRoot\modules\TreeDialogue.psm1" -Force
Import-Module "$PSScriptRoot\modules\ProgressHelper.psm1" -Force
Import-Module "$PSScriptRoot\modules\ExcelHelper.psm1" -Force
Import-Module "$PSScriptRoot\modules\PdfHelper.psm1" -Force
Import-Module "$PSScriptRoot\modules\HelperFunctions.psm1" -Force

$wordExtensions = @(".doc", ".dot", ".wbk", ".docx", ".docm", ".dotx", ".dotm", ".docb")
$pdfExtensions = @(".pdf")

$allExtensions = ($wordExtensions + $pdfExtensions)

# enforce template-description-file to be passed as parameter
if (-not $PSBoundParameters.ContainsKey('template-description-file')) {
    exitError "Eine Vorlagenbeschreibungsdatei muss als Parameter übergeben werden!"
}

$useSelectedDescriptionFile = -not ${no-selected-description-file}.IsPresent
$showTreeSelection = -not ${skip-tree-selection}.IsPresent
if (-not $showTreeSelection) {
    $useSelectedDescriptionFile = $false
}

# make sure working directory exists and make path always absolute
$workingDirectory = makePathAbsolute (Get-Location).Path ${working-directory}
if (-not (Test-Path $workingDirectory)) {
    exitError "Das Arbeitsverzeichnis existiert nicht: $_"
}

# get replacement variables from Excel table
if (${get-variables-from-excel}.IsPresent) {
    # base relative paths upon working directory
    $excelWorkbookFile = makePathAbsolute $workingDirectory ${excel-variables-workbook-file}
    
    if (-not (Test-Path $excelWorkbookFile)) {
        exitError "Die Excel-Arbeitsmappe $excelWorkbookFile existiert nicht!"
    }
        
    try {
        $variables += [string[][]](Get-ExcelTable $excelWorkbookFile ${excel-variables-table-name} ${excel-variables-worksheet-name} | ForEach-Object {@(, ($_.PSObject.Properties | ForEach-Object {$_.Value}))})
    } catch {
        exitError "Beim Auslesen der Tabelle mit den Variablen ist ein Fehler aufgetreten: $($_.Exception.Message)`nStimmen der Arbeitsmappen-, Arbeitsblatt- und Tabellenname???"
    }
}

# get translations from Excel table stuff
$translationHeadings = @()
$translationHeadingsInExcel = ${get-translations-from-excel}.IsPresent
if ($translationHeadingsInExcel) {
    # base relative paths upon working directory
    $excelWorkbookFile = makePathAbsolute $workingDirectory ${excel-translations-workbook-file}
    
    if (-not (Test-Path $excelWorkbookFile)) {
        exitError "Die Excel-Arbeitsmappe $excelWorkbookFile existiert nicht!"
    }
        
    try {
        $gottenTranslations = [string[][]](Get-ExcelTable $excelWorkbookFile ${excel-translations-table-name} ${excel-translations-worksheet-name} | ForEach-Object {@(, ($_.PSObject.Properties | ForEach-Object {$_.Value}))})
        $translationHeadings += $gottenTranslations | ForEach-Object {,@(('^' + [regex]::Escape($_[0]) + '$'), $_[1])}
    } catch {
        exitError "Beim Auslesen der Tabelle mit den Übersetzungen ist ein Fehler aufgetreten: $($_.Exception.Message)`nStimmen der Arbeitsmappen-, Arbeitsblatt- und Tabellenname???"
    }
}

# check if custom template pdf page can be used if given as parameter
if ($PSBoundParameters.ContainsKey('custom-template-pdf-page')) {
    $Script:customTemplatePdfPage = makePathAbsolute $workingDirectory ${custom-template-pdf-page}
    try {
        $extension = (Get-Item $customTemplatePdfPage -ErrorAction Stop).Extension
        if (-not $wordExtensions.Contains($extension.ToLower())) { throw }
    } catch {
        exitError "Es wurde der Pfad zu einer benutzerdefinierten PDF-Vorlagenseite als Parameter übergeben, aber die Datei existiert nicht oder ist keine Word-Datei!"
    }
}

# ask if existing target document should be deleted or check if its base path exists
$targetFile = makePathAbsolute $workingDirectory ${target-file}

try {
    if (Test-Path $targetFile) {
        if ((yesNoBox 'Zieldokument existiert bereits' "Soll $targetFile überschrieben werden?" 'No' 'Warning') -eq 'No') {
            throw "Abgebrochen"
        }
        Remove-Item $targetFile -ErrorAction Stop
    } elseif (-not (Split-Path $targetFile | Test-Path)) {
        throw "Das Verzeichnis, in das das Zieldokument soll, existiert nicht!"
    }
} catch {
    exitError $_.Exception.Message
}

# base selected description file name and path upon target file if not explicitely passed
$selectedDescriptionFile = makePathAbsolute $workingDirectory ([System.IO.Path]::ChangeExtension(${target-file}, "desc"))
if ($PSBoundParameters.ContainsKey('selected-description-file')) {
    $selectedDescriptionFile = makePathAbsolute $workingDirectory ${selected-description-file}

    if (-not (Split-Path $selectedDescriptionFile | Test-Path)) {
        exitError "Das Verzeichnis, in dem die ausgewählte Beschreibungsdatei gespeichert werden soll, existiert nicht!"
    }
}

# check if template description file exists and make path always absolute
try {
    $Script:templateDescriptionFile = Resolve-Path (makePathAbsolute $workingDirectory ${template-description-file}) -ErrorAction Stop
} catch {
    exitError "Die Beschreibungsdatei der Vorlagen existiert nicht: $_"
}

# ask if existing selected description should be used or check if selected description's base path exists
$descriptionFile = $templateDescriptionFile
if ($useSelectedDescriptionFile) {
    if (Test-Path $selectedDescriptionFile) {
        if ((yesNoBox 'Auswahl bereits getroffen' "Soll die bereits getroffene Auswahl verwendet werden?`nWenn nicht, wird sie überschrieben.") -eq 'Yes') {
            $descriptionFile = $selectedDescriptionFile
        }
    } elseif (-not (Split-Path $selectedDescriptionFile | Test-Path)) {
        exitError "Das Verzeichnis, in dem die Datei mit der getroffenen Auswahl gespeichert werden soll, existiert nicht!"
    }
}

# extract description from file
try {
    $Script:description = getDescription $descriptionFile
} catch {
    exitError "Die Beschreibungsdatei konnte nicht ausgelesen werden: $($_.Exception.Message)"
}

Write-Debug "Description before tree selection:"
$description | Select-Object -Property * -ExcludeProperty rawflags,asset | Format-Table | Out-String | Write-Debug

if ($showTreeSelection) {
    $continue = showTree $description
    if ($continue -ne $true) {
        exitError "Abgebrochen"
    }
}

Write-Debug "Description after tree selection:"
$description | Select-Object -Property * -ExcludeProperty rawflags,asset | Format-Table | Out-String | Write-Debug

# write selected elements of description to file
if ($useSelectedDescriptionFile) {
    setDescription $selectedDescriptionFile $description
}

$description = ($description | Where-Object { $_.enabled })
$totalOperations = 6
foreach ($d in $description) { $totalOperations++ }

# custom base paths
$customBasePaths = ${custom-base-path}
Write-Debug "custom base paths:"
$customBasePaths | Format-List | Out-String | Write-Debug

$WA = WordAbstraction
if ($PSBoundParameters.ContainsKey('custom-template-pdf-page')) {
    $WA.templatePdfPage = $customTemplatePdfPage
}

$progress = ProgressHelper "Generiere Word-Dokument ..."
$progress.setTotalOperations($totalOperations)

try {
    $customBasePathRecursion = [System.Collections.Stack]@()

    $addPageBreakInBetweenNextIndent = [System.Collections.Stack]@()

    $pdfHeadings = [System.Collections.ArrayList]@()
    $pdfRecursiveHeading = $false
    $pdfRecursiveHeadingStartIndent = 0

    foreach ($d in $description) {
        $progress.update("Hänge $($d.desc) an")

        $pieces = [System.Collections.Queue]@()
        $pieces.Enqueue($d) | Out-Null
        
        while ($pieces.Count -gt 0) {
            $p = $pieces.Dequeue()
            $p | Format-List

            ### custom base path stuff
            # pop customBasePathRecursions from stack (loop because indent could jump down more than one)
            while ($customBasePathRecursion.Count -gt 0 -and $p.indent -le $customBasePathRecursion.Peek().startIndent) {
                $customBasePathRecursion.Pop() | Out-Null
            }

            # set template path (relative to description file if custom base path flag not set)
            $basePath = Split-Path $templateDescriptionFile
            if ($p.flags.ContainsKey("customBasePath")) {
                if ($p.flags.customBasePath -lt 0 -or $p.flags.customBasePath -ge $customBasePaths.Count) {
                    throw "Kein $($p.flags.customBasePath + 1). benutzerdefinierter Basispfad als Parameter übergeben!"
                }
    
                $basePath = makePathAbsolute $workingDirectory $customBasePaths[$p.flags.customBasePath]

                # check if it should recurse and add to stack
                if ($p.flags.ContainsKey("customBasePathRecursion")) {
                    $customBasePathRecursion.Push(@{
                        basePath = $basePath
                        startIndent = $p.indent
                    }) | Out-Null
                }
            } elseif ($customBasePathRecursion.Count -gt 0) {
                $basePath = $customBasePathRecursion.Peek().basePath
            }
    
            # append language folder if not skipped
            if (-not $p.flags.ContainsKey("skipLang") -and $PSBoundParameters.ContainsKey('lang')) {
                $basePath = Join-Path -Path $basePath -ChildPath $lang
            }
    
            # append path from description if it's relative otherwise use it directly
            $p.path = makePathAbsolute $basePath $p.path
            
            ### grab alphabetically sorted items from given folder if flag set
            if ($p.flags.ContainsKey("alphabetical")) {
                if (-not (Get-Item $p.path -ErrorAction Stop) -is [System.IO.DirectoryInfo]) {
                    throw "$($p.path) ist kein Verzeichnis!"
                }
                
                $gottenFiles = (getDescribedFolder -Path $p.path -Recurse:($p.flags.ContainsKey("alphabeticalRecursion")) -Extensions $allExtensions -Indent ($p.indent + ($p.flags.ContainsKey("alphabeticalKeepIndent") ? 0 : 1)))
                foreach ($s in $gottenFiles) {
                    $pieces.Enqueue($s) | Out-Null
                }
            }

            ### page break in between stuff
            $addPageBreakInBetween = $false

            # pop addPageBreakInBetweenNextIndents from stack (loop because indent could jump down more than one)
            while ($addPageBreakInBetweenNextIndent.Count -gt 0 -and $p.indent -le $addPageBreakInBetweenNextIndent.Peek().startIndent) {
                $addPageBreakInBetweenNextIndent.Pop() | Out-Null
            }

            # add page break in between on next indent (pn) or add it before this element (pt)
            if ($p.flags.ContainsKey("pageBreakInBetween")) {
                if ($p.flags.pageBreakInBetween -eq "t") {
                    $addPageBreakInBetween = $true
                } elseif ($p.flags.pageBreakInBetween -eq "n") {
                    $addPageBreakInBetweenNextIndent.Push(@{
                        startIndent = $p.indent
                        firstOnIndent = $true
                    }) | Out-Null
                }
            }

            ### pdf heading stuff
            # clear headings from empty folders
            while ($pdfHeadings.Count -gt 0 -and $p.indent -le $pdfHeadings[$pdfHeadings.Count-1].onIndent) {
                $pdfHeadings.RemoveAt($pdfHeadings.Count-1) | Out-Null
            }

            # find out pdf heading tier
            $pdfHeadingTier = "None"
            if ($p.indent -le $pdfRecursiveHeadingStartIndent) {
                $pdfRecursiveHeading = $false
            }
            if ($p.flags.ContainsKey("headingTier")) {
                if ($p.flags.headingTier -eq "r") {
                    if (-not $pdfRecursiveHeading) {
                        $pdfRecursiveHeading = $true
                        $pdfRecursiveHeadingStartIndent = $p.indent
                        $pdfHeadingTier = $p.indent+1
                    }
                } else {
                    $pdfHeadingTier = $p.flags.headingTier
                }
            } elseif ($pdfRecursiveHeading) {
                $pdfHeadingTier = $p.indent+1
            }

            # find out pdf heading text
            $pdfHeadingText = $p.desc
            foreach ($translation in $translationHeadings) {
                if ($pdfHeadingText -match $translation[0]) {
                    $pdfHeadingText = $pdfHeadingText -replace $translation
                    break
                }
            }
            
            ### if set only description, nothing to concatenate
            if ($p.flags.ContainsKey("descOnly")) {
                $pdfHeadings.Add(@{
                    pdfHeadingTier = $pdfHeadingTier
                    pdfHeadingText = $pdfHeadingText
                    onIndent = $p.indent
                }) | Out-Null

                continue
            }

            ### concatenation stuff

            # check if file exists while retrieving file type
            $extension = (Get-Item $p.path -ErrorAction Stop).Extension

            if ($wordExtensions.Contains($extension.ToLower())) {
                # only add page break in between --> not on first element on same indentation
                if ($addPageBreakInBetweenNextIndent.Count -gt 0 -and $addPageBreakInBetweenNextIndent.Peek().startIndent + 1 -eq $p.indent) {
                    if ($addPageBreakInBetweenNextIndent.Peek().firstOnIndent) {
                        $addPageBreakInBetweenNextIndent.Peek().firstOnIndent = $false
                    } else {
                        $addPageBreakInBetween = $true
                    }
                }
                if (-not $WA.concatenate($targetFile, $p.path, $addPageBreakInBetween)) { $progress.error() }
                $pdfHeadings.Clear()
            } elseif ($pdfExtensions.Contains($extension.ToLower())) {
                $nPages = getPdfPageNumber($p.path)
                Write-Debug "pdf page number: $nPages"
                for ($i = 1; $i -le $nPages; $i++) {
                    $pdfHeadings.Add(@{
                        pdfHeadingTier = $pdfHeadingTier
                        pdfHeadingText = $pdfHeadingText
                        onIndent = $p.indent
                    }) | Out-Null
                    if (-not $WA.concatenatePdfPage($targetFile, $p.path, $i, $pdfHeadings)) { $progress.error() }
                    $pdfHeadings.Clear()
                    $pdfHeadingTier = "None"
                }
            } else {
                throw "Der Dateityp der Datei $($p.path) ist nicht unterstützt!"
            }
        }
    }
 
    $progress.update("Variablen ersetzen")
    foreach ($variable in $variables) {
        [string]$name = $variable[0]
        [string]$value = $variable[1]

        if ($name -eq "") { continue }
        
        if (-not $WA.replaceVariable($targetFile, $name, $value)) { $progress.error() }
    }

    $progress.update("Aktualisiere Überschriften")
    if (-not $WA.updateHeadings($targetFile)) { $progress.error() }
    
    $progress.update("Aktualisiere Felder")
    if (-not $WA.updateFields($targetFile)) { $progress.error() }
    
    $progress.update("Wende Formatierungsmakierungen an")
    if (-not $WA.applyFormatting($targetFile)) { $progress.error() }
    
    $progress.update("Speichern und schließen")
    if (-not $WA.saveAndClose($targetFile)) { $progress.error() }

    $progress.success()
} catch {
    try { $WA.saveAndClose($targetFile) | Out-Null } catch {}
    $_ | Out-String | Write-Error
    exitError "Zusammensetzen leider fehlgeschlagen: $($_.Exception.Message)"
} finally {
    $WA.destroy()
    $progress.finish()
}

exit 0
