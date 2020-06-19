Param (
    #[parameter(Mandatory=$true)]
    [String] ${working-directory} = "X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch",

    [String] ${target-file} = ".\Ziel.docx",
    [String] ${desciption-file} = "X:\Vorlagen\Bedienhandbuch\Vorlage.desc",
    [String] $lang = "de",
    [String] ${variable-replacement-excel-range}
)
$PSBoundParameters.ContainsKey("variable-replacement-excel-range")
Import-Module "$PSScriptRoot\modules\WordAbstraction.psm1" -Force
Import-Module "$PSScriptRoot\modules\DescriptionFile.psm1" -Force
Import-Module "$PSScriptRoot\modules\TreeDialogue.psm1" -Force
Import-Module "$PSScriptRoot\modules\ProgressHelper.psm1" -Force

try {
    if (Test-Path $target) { 
        Remove-Item $target -ErrorAction Stop
    }
} catch {
    Write-Error $_.Exception.Message
    exit -1
}

try {
    $Script:description = getDescription($descriptionPath)
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
$totalOperations = 4
foreach ($d in $description) { $totalOperations++ }

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