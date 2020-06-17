$docmPath = "$PSScriptRoot\Makros.docm"
$module = "xToDoc"

$Word = $Null
$docm = $Null

Function initWA() {
    $Script:Word = New-Object -ComObject Word.Application
    $Script:Word.Visible = $False
    $Script:docm = $Word.Documents.Open($docmPath, $False, $True)

    return $Script:Word
}

Function destroyWA() {
    $Script:docm.close($False)
    $Script:Word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
}