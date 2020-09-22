Param (
    [String] ${working-directory} = ".",
    
    [String] ${source-word-file} = ".\target.docx",
    [String] ${target-pdf-file}
)

Import-Module "$PSScriptRoot\modules\ProgressHelper.psm1" -Force
Import-Module "$PSScriptRoot\modules\PdfHelper.psm1" -Force
Import-Module "$PSScriptRoot\modules\WordPdfExportHelper.psm1" -Force
Import-Module "$PSScriptRoot\modules\HelperFunctions.psm1" -Force

# make sure working directory exists and make path always absolute
$workingDirectory = makePathAbsolute (Get-Location).Path ${working-directory}
if (-not (Test-Path $workingDirectory)) {
    exitError "Das Arbeitsverzeichnis existiert nicht: $_"
}

# base source file upon working directory
$sourceWordFile = makePathAbsolute $workingDirectory ${source-word-file}
if (-not (Test-Path $sourceWordFile)) {
    exitError "Das Quelldokument $sourceWordFile existiert nicht!"
}

# base target file upon working directory from 
$targetPdfFile = makePathAbsolute $workingDirectory ([System.IO.Path]::ChangeExtension(${source-word-file}, "pdf"))
if ($PSBoundParameters.ContainsKey('target-pdf-file')) {
    $targetPdfFile = makePathAbsolute $workingDirectory ${target-pdf-file}

    if (-not (Split-Path $targetPdfFile | Test-Path)) {
        exitError "Das Verzeichnis, in dem die PDF-Datei gespeichert werden soll, existiert nicht!"
    }
}

$wpeh = WordPdfExportHelper $sourceWordFile

$progress = ProgressHelper "Exportiere Word-Dokument als PDF ..."
$progress.setTotalOperations(7)

try {
    $progress.update("Bestimme Ersatzseiten")
    $replacements = $wpeh.getPdfReplacementPages()
    $replacements | Format-List | Out-String | Write-Debug
    # one group for each pdf file (more efficient --> faster)
    $replacementGroups = $replacements | Group-Object -Property path

    $progress.update("Bestimme Word-Seitenorientierung")
    $targetIsPortrait = $wpeh.isPortrait()

    $progress.update("Verstecke Platzhalter")
    $wpeh.hidePlaceholders()
    
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
