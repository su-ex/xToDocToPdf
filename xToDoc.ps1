Param (
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

# make sure working directory exists and make path always absolute
try {
    $Script:workingDirectory = Resolve-Path ${working-directory} -ErrorAction Stop
} catch {
    Write-Error "Das Arbeitsverzeichnis existiert nicht: $_"
    exit -1
}

# base relative paths upon working directory
$targetFile = ${target-file}
if (-not [System.IO.Path]::IsPathRooted($targetFile)) {
    $targetFile = [System.IO.Path]::GetFullPath((Join-Path -Path $workingDirectory -ChildPath $targetFile))
}
$selectedDescriptionFile = ${selected-description-file}
if (-not [System.IO.Path]::IsPathRooted($selectedDescriptionFile)) {
    $selectedDescriptionFile = [System.IO.Path]::GetFullPath((Join-Path -Path $workingDirectory -ChildPath $selectedDescriptionFile))
}

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
    $excelWorkbookFile = ${excel-workbook-file}
    if (-not [System.IO.Path]::IsPathRooted($excelWorkbookFile)) {
        $excelWorkbookFile = [System.IO.Path]::GetFullPath((Join-Path -Path $workingDirectory -ChildPath $excelWorkbookFile))
    }
    
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
$description | Format-Table | Out-String | Write-Debug

$continue = showTree $description
if ($continue -ne $true) {
    Write-Error "Abgebrochen"
    exit -1
}

Write-Debug "Description after tree selection:"
$description | Format-Table | Out-String | Write-Debug

# write selected elements of description to file
setDescription $selectedDescriptionFile $description

$description = ($description | Where-Object {$_.enabled})
$totalOperations = 3
if ($replaceVariables) { $totalOperations++ }
foreach ($d in $description) { $totalOperations++ }

# set template path
$templatePath = Split-Path $templateDescriptionFile

# custom base paths
$customBasePaths = ${-custom-base-path}
Write-Debug "custom base paths: $customBasePaths"

$WA = WordAbstraction

$progress = ProgressHelper("Generiere Word-Dokument ...")
$progress.setTotalOperations($totalOperations)

try {
    foreach ($d in $description) {
        $progress.update("Hänge $($d.desc) an")

        # ToDo: handle flags

        $path = Join-Path -Path $templatePath -ChildPath $lang
        $path = Join-Path -Path $path -ChildPath $d.path

        $pdfHeadingTier = "1"
        $pdfHeadingText = $d.desc

        # check if file exists while retrieving file type
        $extension = (Get-Item $path -ErrorAction Stop).Extension

        if (@(".doc", ".dot", ".wbk", ".docx", ".docm", ".dotx", ".dotm", ".docb").Contains($extension.ToLower())) {
            if (-not $WA.concatenate($targetFile, $path)) { $progress.error() }
        } elseif ($extension -ieq ".pdf") {
            $nPages = getPdfPageNumber($path)
            Write-Debug "pdf page number: $nPages"
            for ($i = 1; $i -le $nPages; $i++) {
                if (-not $WA.concatenatePdfPage($targetFile, $path, $i, $pdfHeadingTier, $pdfHeadingText)) { $progress.error() }
                $pdfHeadingTier = "None"
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
    Write-Error "Zusammensetzen leider fehlgeschlagen: $($_.Exception.Message)`n`nIm kommenden Word Dialog NICHT `"Abbrechen`" auswählen, wenn man es noch einmal versuchen will, da man sonst Word per Task Manager beenden muss!!!"
}

$WA.destroy()

exit 0
