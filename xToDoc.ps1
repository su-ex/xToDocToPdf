﻿Param (
    [String] ${working-directory} = "X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch",
    
    [String] ${target-file} = ".\Ziel.docx",
    [String] ${selected-description-file} = ".\Auswahl.desc",

    [String] ${template-description-file} = "X:\Vorlagen\Bedienhandbuch\Vorlage.desc",

    [String] $lang = "de",
    
    [Switch] ${get-variables-from-excel},
    [String] ${excel-workbook-file},
    [String] ${excel-worksheet-name},
    [String] ${excel-table-name},

    [String[]] ${custom-base-path} = @(".")
)

if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $Script:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
} else {
    $Script:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath) { $Script:ScriptPath = "." }
}

Import-Module "$ScriptPath\modules\WordAbstraction.psm1" -Force
Import-Module "$ScriptPath\modules\DescriptionFile.psm1" -Force
Import-Module "$ScriptPath\modules\TreeDialogue.psm1" -Force
Import-Module "$ScriptPath\modules\ProgressHelper.psm1" -Force
Import-Module "$ScriptPath\modules\ExcelHelper.psm1" -Force
Import-Module "$ScriptPath\modules\PdfHelper.psm1" -Force
Import-Module "$ScriptPath\modules\HelperFunctions.psm1" -Force

$wordExtensions = @(".doc", ".dot", ".wbk", ".docx", ".docm", ".dotx", ".dotm", ".docb")
$pdfExtensions = @(".pdf")

# make sure working directory exists and make path always absolute
try {
    $Script:workingDirectory = Resolve-Path ${working-directory} -ErrorAction Stop
} catch {
    exitError "Das Arbeitsverzeichnis existiert nicht: $_"
}

# base relative paths upon working directory
$targetFile = makePathAbsolute $workingDirectory ${target-file}
$selectedDescriptionFile = makePathAbsolute $workingDirectory ${selected-description-file}

# ask if existing target document should be deleted or check if its base path exists
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

# check if template description file exists and make path always absolute
try {
    $Script:templateDescriptionFile = Resolve-Path (makePathAbsolute $workingDirectory ${template-description-file}) -ErrorAction Stop
} catch {
    exitError "Die Beschreibungsdatei der Vorlagen existiert nicht: $_"
}

# ask if existing selected description should be used or check if selected description's base path exists
$descriptionFile = $templateDescriptionFile
if (Test-Path $selectedDescriptionFile) {
    if ((yesNoBox 'Auswahl bereits getroffen' "Soll die bereits getroffene Auswahl verwendet werden?`nWenn nicht, wird sie überschrieben.") -eq 'Yes') {
        $descriptionFile = $selectedDescriptionFile
    }
} elseif (-not (Split-Path $selectedDescriptionFile | Test-Path)) {
    exitError "Das Verzeichnis, in dem die Datei mit der getroffenen Auswahl gespeichert werden soll, existiert nicht!"
}

# replacement variables stuff
$replaceVariables = [boolean]${get-variables-from-excel}
if ($replaceVariables) {
    # base relative paths upon working directory
    $excelWorkbookFile = makePathAbsolute $workingDirectory ${excel-workbook-file}
    
    if (-not (Test-Path $excelWorkbookFile)) {
        exitError "Die Excel-Arbeitsmappe $excelWorkbookFile existiert nicht!"
    }
        
    try {
        $Script:replacementVariables = Get-ExcelTable $excelWorkbookFile ${excel-table-name} ${excel-worksheet-name}
    } catch {
        exitError "Beim Auslesen der Tabelle mit den Variablen ist ein Fehler aufgetreten: $($_.Exception.Message)`nStimmt z. B. der Tabellenname???"
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

$continue = showTree $description
if ($continue -ne $true) {
    exitError "Abgebrochen"
}

Write-Debug "Description after tree selection:"
$description | Select-Object -Property * -ExcludeProperty rawflags,asset | Format-Table | Out-String | Write-Debug

# write selected elements of description to file
setDescription $selectedDescriptionFile $description

$description = ($description | Where-Object {$_.enabled -and $_.path -ne ""})
$totalOperations = 3
if ($replaceVariables) { $totalOperations++ }
foreach ($d in $description) { $totalOperations++ }

# custom base paths
$customBasePaths = ${custom-base-path}
Write-Debug "custom base paths:"
$customBasePaths | Format-List | Out-String | Write-Debug

$WA = WordAbstraction

$progress = ProgressHelper "Generiere Word-Dokument ..."
$progress.setTotalOperations($totalOperations)

try {
    foreach ($d in $description) {
        $progress.update("Hänge $($d.desc) an")
        # Write-Debug "current to concatenate:"
        # $d | Format-List | Out-String | Write-Debug

        # set template path (relative to description file if custom base path flag not set)
        $basePath = Split-Path $templateDescriptionFile
        if ($d.flags.ContainsKey("customBasePath")) {
            if ($d.flags.customBasePath -lt 0 -or $d.flags.customBasePath -ge $customBasePaths.Count) {
                throw "Kein $($d.flags.customBasePath + 1). benutzerdefinierter Basispfad als Parameter übergeben!"
            }

            $basePath = makePathAbsolute $workingDirectory $customBasePaths[$d.flags.customBasePath]
        }

        # append language folder if not skipped
        if (-not $d.flags.ContainsKey("skipLang")) {
            $basePath = Join-Path -Path $basePath -ChildPath $lang
        }

        # append path from description if it's relative otherwise use it directly
        $d.path = makePathAbsolute $basePath $d.path

        [System.Collections.Queue]$pieces = @()
        $pieces = New-Object System.Collections.Queue  
        $pieces.Enqueue($d) | Out-Null
        $pieces | Format-Table
        $pieces.Count
        while ($pieces.Count -gt 0) {
            Write-Host "hier"
            $p = $pieces.Dequeue()
            Write-Host "hier2"
            $p | Format-List

            $pdfHeadingTier = "None"
                ## ToDo: add to list, later remove from list on first pdf file
            $pdfHeadingText = $p.desc
            if ($p.flags.ContainsKey("headingTier")) {
                $pdfHeadingTier = $p.flags.headingTier
                $pdfHeadingText = $p.desc
            }
            
            if ($p.flags.ContainsKey("alphabetical")) {
                if (-not (Get-Item $p.path) -is [System.IO.DirectoryInfo]) {
                    throw "$($p.path) ist kein Verzeichnis!"
                }
                
                $gottenFiles = (getDescribedFolder -Path $p.path -Recurse:($p.flags.alphabetical) -Extensions ($wordExtensions + $pdfExtensions) -Indent ($p.indent + 1))
                foreach ($s in $gottenFiles) {
                    $pieces.Enqueue($s) | Out-Null
                }

                continue
            }
            $p | Format-List
            # check if file exists while retrieving file type
            $extension = (Get-Item $p.path -ErrorAction Stop).Extension

            if ($wordExtensions.Contains($extension.ToLower())) {
                if (-not $WA.concatenate($targetFile, $p.path)) { $progress.error() }
            } elseif ($pdfExtensions.Contains($extension.ToLower())) {
                $nPages = getPdfPageNumber($p.path)
                Write-Debug "pdf page number: $nPages"
                for ($i = 1; $i -le $nPages; $i++) {
                    if (-not $WA.concatenatePdfPage($targetFile, $p.path, $i, $pdfHeadingTier, $pdfHeadingText)) { $progress.error() }
                    $pdfHeadingTier = "None"
                }
            } else {
                throw "Dateityp nicht unterstützt!"
            }
        }
    }

    if ($replaceVariables) { 
        $progress.update("Variablen ersetzen")
        
        foreach ($variable in $replacementVariables) {
            $name = $variable.Variable

            # split lines and later join them with a vertical tab (Word's new line in same paragraph: Shift+Enter)
            $replacementLines = $variable.Wert -split "`r?`n"

            # get flags from Excel to apply special behaviour on some variables:
            #   "t" to append a horizontal tab before each line in a variable
            $flags = [char[]]$variable.Flags
            foreach ($flag in $flags) {
                if ($flag -eq 't') {
                    $replacementLines = $replacementLines | ForEach-Object { "`t" + $_ }
                }
            }
            
            if (-not $WA.replaceVariable($targetFile, $name, $replacementLines -join "`v")) { $progress.error() }
        }
    }

    $progress.update("Aktualisiere Überschriften")
    if (-not $WA.updateHeadings($targetFile)) { $progress.error() }
    
    $progress.update("Aktualisiere Felder")
    if (-not $WA.updateFields($targetFile)) { $progress.error() }
    
    $progress.update("Speichern und schließen")
    if (-not $WA.saveAndClose($targetFile)) { $progress.error() }

    $progress.success()
} catch {
    try { $WA.saveAndClose($targetFile) | Out-Null } catch {}
    exitError "Zusammensetzen leider fehlgeschlagen: $($_.Exception.Message)"
} finally {
    $WA.destroy()
    $progress.finish()
}

exit 0
