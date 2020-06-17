Import-Module "$PSScriptRoot\modules\WordAbstraction.psm1" -Force
Import-Module "$PSScriptRoot\modules\DescriptionFile.psm1" -Force
Import-Module "$PSScriptRoot\modules\TreeDialogue.psm1" -Force

$description = getDescription("X:\Vorlagen\Bedienhandbuch\Vorlage.desc")
$description | Format-Table

showTree($description)

$description | Format-Table

$path = "X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch"
$target = "$path\Ziel.docx"

$functions = [System.Collections.Queue]@()

$functions.Enqueue({ $Script:WA = initWA })
$functions.Enqueue({ Copy-Item "$path\FormatvorlagenUndAnfang.docx" -Destination "$target" -Force })
$functions.Enqueue({ $WA.Run("concatenate", [ref]$target, [ref]"$path\WichtigeInformation.doc") })
$functions.Enqueue({ $WA.Run("concatenate", [ref]$target, [ref]"$path\blablabla.doc") })
$functions.Enqueue({ $WA.Run("concatenate", [ref]$target, [ref]"$path\blablabla.doc") })
$functions.Enqueue({ $WA.Run("concatenate", [ref]$target, [ref]"$path\Rest.doc") })
$functions.Enqueue({ $WA.Run("updateHeadings", [ref]$target) })
$functions.Enqueue({ $WA.Run("updateFields", [ref]$target) })
$functions.Enqueue({ $WA.Run("saveAndClose", [ref]$target) })
$functions.Enqueue({ destroyWA })

$total = $functions.Count
$i = 0
while ($functions.Count -gt 0) {
    Write-Progress -Activity "Generiere Word-Dokument ..." -Status "$($i+1) / $total" -PercentComplete (($i++/$total)*100) -CurrentOperation "Tue irgendwas"
    & $functions.Dequeue()
}