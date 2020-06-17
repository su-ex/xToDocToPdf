Import-Module "$PSScriptRoot\modules\WordAbstraction.psm1" -Force
Import-Module "$PSScriptRoot\modules\DescriptionFile.psm1" -Force
Import-Module "$PSScriptRoot\modules\TreeDialogue.psm1" -Force

$description = getDescription("X:\Vorlagen\Bedienhandbuch\Vorlage.desc")
$description | Format-Table

showTree($description)

$path = "X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch"
$target = "$path\Ziel.docx"

$WA = initWA

Copy-Item "$path\FormatvorlagenUndAnfang.docx" -Destination "$target" -Force
$WA.Run("concatenate", [ref]$target, [ref]"$path\WichtigeInformation.doc")
$WA.Run("concatenate", [ref]$target, [ref]"$path\blablabla.doc")
$WA.Run("concatenate", [ref]$target, [ref]"$path\blablabla.doc")
$WA.Run("concatenate", [ref]$target, [ref]"$path\Rest.doc")
$WA.Run("updateHeadings", [ref]$target)
$WA.Run("updateFields", [ref]$target)
$WA.Run("saveAndClose", [ref]$target)

destroyWA
