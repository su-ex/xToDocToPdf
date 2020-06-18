#Param ($workingDirectory, $envname)

Import-Module "$PSScriptRoot\modules\WordAbstraction.psm1" -Force
Import-Module "$PSScriptRoot\modules\DescriptionFile.psm1" -Force
Import-Module "$PSScriptRoot\modules\TreeDialogue.psm1" -Force
Import-Module "$PSScriptRoot\modules\JobHandling.psm1" -Force

#$workingDirectory

$descriptionPath = "X:\Vorlagen\Bedienhandbuch\Vorlage.desc"
$lang = "de"
$target = "X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch\Ziel.docx"

if (Test-Path $target) { Remove-Item $target }

$description = getDescription($descriptionPath)
#$description | Format-Table

$continue = showTree($description)
if ($continue -ne $true) {
    Write-Error "Abgebrochen"
    exit -1
}

#$description | Format-Table

$path = Split-Path $descriptionPath
$path = Join-Path -Path $path -ChildPath $lang

$WA = WordAbstraction

$jobs = JobHandling("Generiere Word-Dokument ...")

foreach ($d in $description) {
    if($d.enabled -eq $False) { continue }
    "$target, $(Join-Path -Path $path -ChildPath $d.path)"
    Invoke-Command { $WA.concatenate($target, (Join-Path -Path $path -ChildPath $d.path)) }
}

$jobs.add("Aktualisiere Überschriften", { $WA.updateHeadings($target) })
$jobs.add("Aktualisiere Felder", { $WA.updateFields($target) })
$jobs.add("Speichern", { $WA.saveAndClose($target) })

$jobs.run()

$WA.destroy()

exit 0