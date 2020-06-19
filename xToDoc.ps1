Param (
    [String] ${working-directory} = "X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch",
    
    [String] ${target-file} = ".\Ziel.docx",
    [String] ${selected-description-file} = ".\Auswahl.desc",

    [String] ${template-description-file} = "X:\Vorlagen\Bedienhandbuch\Vorlage.desc",

    [String] $lang = "de",
    
    [Switch] ${get-variables-from-excel},
    [String] ${excel-workbook-file},
    [String] ${excel-worksheet-name},
    [String] ${excel-table-name}
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
        Write-Error "Beim Auslesen der Tabelle mit den Variablen ist ein Fehler aufgetreten: $($_.Exception.Message)`nStimmt der Tabellenname???"
        exit -1
    }
}

try {
    $Script:description = getDescription($descriptionFile)
} catch {
    Write-Error $_.Exception.Message
    exit -1
}
#$description | Format-Table

$continue = showTree($description)
if ($continue -ne $true) {
    Write-Error "Abgebrochen"
    exit -1
}

# ToDo: save description
#$description | Format-Table

$description = ($description | Where {$_.enabled})
#$description | Format-Table
$totalOperations = 3
if ($replaceVariables) { $totalOperations++ }
foreach ($d in $description) { $totalOperations++ }

$templatePath = Split-Path $descriptionFile
$templatePath = Join-Path -Path $templatePath -ChildPath $lang

$WA = WordAbstraction

$progress = ProgressHelper("Generiere Word-Dokument ...")
$progress.setTotalOperations($totalOperations)

try {
    # ToDo: handle non doc
    foreach ($d in $description) {
        $progress.update("Hänge $($d.desc) an")
        if (-not $WA.concatenate($targetFile, (Join-Path -Path $templatePath -ChildPath $d.path)) ) { $progress.error() }
    }

    if ($replaceVariables) { 
        $progress.update("Variablen ersetzen")
        # ToDo: Variablen ersetzen
    }

    $progress.update("Aktualisiere Überschriften")
    if (-not $WA.updateHeadings($targetFile)) { $progress.error() }
    
    $progress.update("Aktualisiere Felder")
    if (-not $WA.updateFields($targetFile)) { $progress.error() }
    
    $progress.update("Speichern")
    if (-not $WA.saveAndClose($targetFile)) { $progress.error() }

    $progress.success()
} catch {
    Write-Error $_.Exception.Message
}

$WA.destroy()

exit 0
