using module .\modules\WordAbstraction.psm1
using module .\modules\DescriptionFile.psm1
using module .\modules\TreeDialogue.psm1
using module .\modules\JobHandling.psm1

$description = getDescription("X:\Vorlagen\Bedienhandbuch\Vorlage.desc")
$description | Format-Table

showTree($description)

$description | Format-Table

$path = "X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch"
$target = "$path\Ziel.docx"

$jobs = [JobHandling]::new("Generiere Word-Dokument ...")

$jobs.add("Tue irgendwas", { $Script:WA = initWA })
$jobs.add("Tue irgendwas", { Copy-Item "$path\FormatvorlagenUndAnfang.docx" -Destination "$target" -Force })
$jobs.add("Tue irgendwas", { $WA.Run("concatenate", [ref]$target, [ref]"$path\WichtigeInformation.doc") })
$jobs.add("Tue irgendwas", { $WA.Run("concatenate", [ref]$target, [ref]"$path\blablabla.doc") })
$jobs.add("Tue irgendwas", { $WA.Run("concatenate", [ref]$target, [ref]"$path\blablabla.doc") })
$jobs.add("Tue irgendwas", { $WA.Run("concatenate", [ref]$target, [ref]"$path\Rest.doc") })
$jobs.add("Tue irgendwas", { $WA.Run("updateHeadings", [ref]$target) })
$jobs.add("Tue irgendwas", { $WA.Run("updateFields", [ref]$target) })
$jobs.add("Tue irgendwas", { $WA.Run("saveAndClose", [ref]$target) })
$jobs.add("Tue irgendwas", { destroyWA })

$jobs.run()