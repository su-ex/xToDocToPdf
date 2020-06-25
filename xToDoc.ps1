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
    Write-Error "Das Arbeitsverzeichnis existiert nicht: $_"
    exit -1
}

# base relative paths upon working directory
$targetFile = makePathAbsolute $workingDirectory ${target-file}
$selectedDescriptionFile = makePathAbsolute $workingDirectory ${selected-description-file}

# ask if existing target document should be deleted or check if its base path exists
try {
    if (Test-Path $targetFile) {
        if ($host.ui.PromptForChoice('Zieldokument existiert bereits', "Soll $targetFile überschrieben werden?", @("Ja", "Nein"), 1) -eq 1) {
            throw "Abgebrochen"
        }
        Remove-Item $targetFile -ErrorAction Stop
    } elseif (-not (Split-Path $targetFile | Test-Path)) {
        throw "Das Verzeichnis, in das das Zieldokument soll, existiert nicht!"
    }
} catch {
    Write-Error $_.Exception.Message
    exit -1
}

# check if template description file exists and make path always absolute
try {
    $Script:templateDescriptionFile = Resolve-Path ${template-description-file} -ErrorAction Stop
} catch {
    Write-Error "Die Beschreibungsdatei der Vorlagen existiert nicht: $_"
    exit -1
}

# ask if existing selected description should be used or check if selected description's base path exists
$descriptionFile = $templateDescriptionFile
if (Test-Path $selectedDescriptionFile) {
    if ($host.ui.PromptForChoice('Auswahl bereits getroffen', "Soll die bereits getroffene Auswahl verwendet werden?`nWenn nicht, wird sie überschrieben.", @("Ja", "Nein"), 0) -eq 0) {
        $descriptionFile = $selectedDescriptionFile
    }
} elseif (-not (Split-Path $selectedDescriptionFile | Test-Path)) {
    Write-Error "Das Verzeichnis, in dem die Datei mit der getroffenen Auswahl gespeichert werden soll, existiert nicht!"
    exit -1
}

# replacement variables stuff
$replaceVariables = [boolean]${get-variables-from-excel}
if ($replaceVariables) {
    # base relative paths upon working directory
    $excelWorkbookFile = makePathAbsolute $workingDirectory ${excel-workbook-file}
    
    if (-not (Test-Path $excelWorkbookFile)) {
        Write-Error "Die Excel-Arbeitsmappe $excelWorkbookFile existiert nicht!"
        exit -1
    }
        
    try {
        $Script:replacementVariables = Get-ExcelTable $excelWorkbookFile ${excel-table-name} ${excel-worksheet-name}
    } catch {
        Write-Error "Beim Auslesen der Tabelle mit den Variablen ist ein Fehler aufgetreten: $($_.Exception.Message)`nStimmt z. B. der Tabellenname???"
        exit -1
    }
}

# extract description from file
try {
    $Script:description = getDescription $descriptionFile
} catch {
    Write-Error "Die Beschreibungsdatei konnte nicht ausgelesen werden: $($_.Exception.Message)"
    exit -1
}

Write-Debug "Description before tree selection:"
$description | Select-Object -Property * -ExcludeProperty raw,asset | Format-Table | Out-String | Write-Debug

$continue = showTree $description
if ($continue -ne $true) {
    Write-Error "Abgebrochen"
    exit -1
}

Write-Debug "Description after tree selection:"
$description | Select-Object -Property * -ExcludeProperty raw,asset | Format-Table | Out-String | Write-Debug

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

$progress = ProgressHelper("Generiere Word-Dokument ...")
$progress.setTotalOperations($totalOperations)

try {
    foreach ($d in $description) {
        $progress.update("Hänge $($d.desc) an")
        # Write-Debug "current to concatenate:"
        # $d | Format-List | Out-String | Write-Debug

        # set template path (relative to description file if custom base path flag not set)
        $path = Split-Path $templateDescriptionFile
        if ($d.flags.ContainsKey("customBasePath")) {
            if ($d.flags.customBasePath -lt 0 -or $d.flags.customBasePath -ge $customBasePaths.Count) {
                throw "Kein $($d.flags.customBasePath + 1). benutzerdefinierter Basispfad als Parameter übergeben!"
            }

            $path = makePathAbsolute $workingDirectory $customBasePaths[$d.flags.customBasePath]
        }

        # append language folder if not skipped
        if (-not $d.flags.ContainsKey("skipLang")) {
            $path = Join-Path -Path $path -ChildPath $lang
        }

        # append path from description if it's relative otherwise use it directly
        $path = makePathAbsolute $path $d.path

        $pieces = [System.Collections.ArrayList]@()

        if ($d.flags.ContainsKey("alphabetical")) {
            if (-not (Get-Item $path) -is [System.IO.DirectoryInfo]) {
                throw "$path ist kein Verzeichnis!"
            }
            
            $gottenFiles = (getDescribedFolder -Path $path -Recurse:($d.flags.alphabetical) -Extensions ($wordExtensions + $pdfExtensions))
            foreach ($s in $gottenFiles) {
                $pieces.Add($s) | Out-Null
            }
        } else {
            $d.path = $path
            $pieces.Add($d) | Out-Null
        }

        $pdfHeadingTier = "None"
        if ($d.flags.ContainsKey("headingTier")) {
            $pdfHeadingTier = $d.flags.headingTier
        }

        foreach ($p in $pieces) {
            if ($p.flags.ContainsKey("headingTier")) {
                $pdfHeadingTier = $p.flags.headingTier
            }
            $pdfHeadingText = ItIf $p.desc $d.desc

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
        # $progress.update("Variablen ersetzen")
        
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
} catch {
    try { $WA.saveAndClose($targetFile) | Out-Null } catch {}
    Write-Error "Zusammensetzen leider fehlgeschlagen: $($_.Exception.Message)"
    exit -1
} finally {
    $WA.destroy()
}

$progress.success()

exit 0
