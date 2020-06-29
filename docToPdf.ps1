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
    Write-Error "Das Arbeitsverzeichnis existiert nicht: $_"
    exit -1
}

# base source file upon working directory
$sourceWordFile = makePathAbsolute $workingDirectory ${source-word-file}
if (-not (Test-Path $sourceWordFile)) {
    Write-Error "Das Quelldokument $sourceWordFile existiert nicht!"
    exit -1
}

# base target file upon working directory from 
$targetPdfFile = makePathAbsolute $workingDirectory ([System.IO.Path]::ChangeExtension(${source-word-file}, "pdf"))
if (${target-pdf-file} -ne "") {
    $targetPdfFile = makePathAbsolute $workingDirectory ${target-pdf-file}

    if (-not (Split-Path $targetPdfFile | Test-Path)) {
        Write-Error "Das Verzeichnis, in dem die PDF-Datei gespeichert werden soll, existiert nicht!"
        exit -1
    }
}

$wpeh = WordPdfExportHelper $sourceWordFile

$progress = ProgressHelper "Exportiere Word-Dokument als PDF ..."
$progress.setTotalOperations(4)

try {
    $progress.update("Bestimme Ersatzseiten")
    $replacements = $wpeh.getPdfReplacementPages()
    $replacements | Out-String | Write-Debug

    $progress.update("Verstecke Platzhalter")
    $wpeh.hidePlaceholders()
    
    $progress.update("Exportiere als PDF")
    $wpeh.export($targetPdfFile)
    
    $progress.update("Überlagere Seiten")
    overlayPdfPages $targetPdfFile $replacements
} catch {
    Write-Error "Export leider fehlgeschlagen: $($_.Exception.Message)"
    exit -1
} finally {
    $wpeh.destroy()
    $progress.finish()
}

exit 0
