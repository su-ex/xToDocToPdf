#Param ($workingDirectory, $envname)

Import-Module "$PSScriptRoot\modules\WordAbstraction.psm1" -Force
Import-Module "$PSScriptRoot\modules\DescriptionFile.psm1" -Force
Import-Module "$PSScriptRoot\modules\TreeDialogue.psm1" -Force
Import-Module "$PSScriptRoot\modules\ProgressHelper.psm1" -Force

#$workingDirectory

$descriptionPath = "X:\Vorlagen\Bedienhandbuch\Vorlage.desc"
$lang = "de"
$target = "X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch\Ziel.docx"

if (Test-Path $target) { Remove-Item $target }

try {
    $Script:description = getDescription($descriptionPath)
} catch {
    Write-Error $_.Exception.Message
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
$totalOperations = $description.Count + 4

$path = Split-Path $descriptionPath
$path = Join-Path -Path $path -ChildPath $lang

$WA = WordAbstraction

$progress = ProgressHelper("Generiere Word-Dokument ...")
$progress.setTotalOperations($totalOperations)
 
try {
    # ToDo: handle non doc
    foreach ($d in $description) {
        $progress.update("Hänge $($d.desc) an")
        if (-not $WA.concatenate($target, (Join-Path -Path $path -ChildPath $d.path)) ) { $progress.error() }
    }

    $progress.update("Variablen ersetzen")
    # ToDo: Variablen ersetzen

    $progress.update("Aktualisiere Überschriften")
    if (-not $WA.updateHeadings($target)) { $progress.error() }
    
    $progress.update("Aktualisiere Felder")
    if (-not $WA.updateFields($target)) { $progress.error() }
    
    $progress.update("Speichern")
    if (-not $WA.saveAndClose($target)) { $progress.error() }

    $progress.success()
} catch {
    Write-Error $_.Exception.Message
}

$WA.destroy()

exit 0