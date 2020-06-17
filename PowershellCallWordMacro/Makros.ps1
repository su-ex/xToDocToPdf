$docmPath = "$PSScriptRoot\Makros.docm"
$module = "xToDoc"

$Word = New-Object -ComObject Word.Application
$Word.Visible = $False
$docm = $Word.Documents.Open($docmPath, $False, $True)

$path = "X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch"
$target = "$path\Ziel.docx"

Copy-Item "$path\FormatvorlagenUndAnfang.docx" -Destination "$target" -Force
$Word.Run("$module.concatenate", [ref]$target, [ref]"$path\WichtigeInformation.doc")
$Word.Run("$module.concatenate", [ref]$target, [ref]"$path\blablabla.doc")
$Word.Run("$module.concatenate", [ref]$target, [ref]"$path\Rest.doc")
$Word.Run("$module.updateHeadings", [ref]$target)
$Word.Run("$module.updateFields", [ref]$target)
$Word.Run("$module.saveAndClose", [ref]$target)

$docm.close($False)
$Word.Quit()
$a=[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word)