Param (
    [String] ${working-directory} = "X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch",
    
    [String] ${source-word-file} = ".\Ziel.docx",
    [String] ${target-pdf-file} = ""
)

if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $Script:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
} else {
    $Script:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath) { $Script:ScriptPath = "." }
}

Import-Module "$ScriptPath\modules\WordAbstraction.psm1" -Force
Import-Module "$ScriptPath\modules\ProgressHelper.psm1" -Force
Import-Module "$ScriptPath\modules\PdfHelper.psm1" -Force
Import-Module "$ScriptPath\modules\WordPdfExportHelper.psm1" -Force
Import-Module "$ScriptPath\modules\HelperFunctions.psm1" -Force

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

# base target file upon working directory from 
$targetPdfFile = makePathAbsolute $workingDirectory ([System.IO.Path]::ChangeExtension(${source-word-file}, "pdf"))
if (${target-pdf-file} -ne "") {
    $targetPdfFile = makePathAbsolute $workingDirectory ${target-pdf-file}

    if (-not (Split-Path $targetPdfFile | Test-Path)) {
        exitError "Das Verzeichnis, in dem die PDF-Datei gespeichert werden soll, existiert nicht!"
    }
}

$wpeh = WordPdfExportHelper $sourceWordFile

$progress = ProgressHelper "Exportiere Word-Dokument als PDF ..."
$progress.setTotalOperations(6)

try {
    $progress.update("Bestimme Ersatzseiten")
    $replacements = $wpeh.getPdfReplacementPages()
    $replacements | Out-String | Write-Debug
    # one group for each pdf file (more efficient --> faster)
    $replacementGroups = $replacements | Group-Object -Property path

    $progress.update("Verstecke Platzhalter")
    $wpeh.hidePlaceholders()

    $targetIsPortrait = $true
    
    $progress.update("Exportiere Word-Dokument als PDF")
    $wpeh.export($targetPdfFile)
    
    $progress.update("Extrahiere PDF-Seitendimensionen und bestimme zu drehende Seiten")
    $targetPagesToRotate = [System.Collections.ArrayList]@()
    foreach ($rg in $replacementGroups) {
        $pageDimensions = getPdfPageDimensions $rg.Name
        foreach ($pd in $pageDimensions) {
            if ($pd.isPortrait -ne $targetIsPortrait) {
                $targetPagesToRotate.AddRange([int[]]($rg.Group | Where-Object { $_.pdfPageNumber -eq $pd.pageNumber } | Select-Object -ExpandProperty docPageNumber))
            }
        }
    }

    $progress.update("Drehe Seiten")
    rotatePdfPages90Deg $targetPdfFile $targetPagesToRotate

    $progress.update("Überlagere Seiten")
    overlayPdfPages $targetPdfFile $replacementGroups

    $progress.success()
} catch {
    $_ | Out-String | Write-Error
    exitError "Export leider fehlgeschlagen: $($_.Exception.Message)"
} finally {
    $wpeh.destroy()
    $progress.finish()
}

exit 0
