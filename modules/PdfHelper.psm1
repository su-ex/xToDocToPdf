$qpdfexe = (Resolve-Path "$PSScriptRoot\..\..\qpdf-**\**\bin\qpdf.exe")

Function getPdfPageNumber($file) {
    #Write-Host "exe: $qpdfexe, file: $file"
    return [int](& "$qpdfexe" --show-npages "`"$file`"")
}

Function overlayPages($targetPdfFile, $replacements) {
    # return [int](& "$qpdfexe" '$targetPdfFile' --overlay 'X:\Vorlagen\Bedienhandbuch\de\SaudiArab60-100.pdf' --from=1-14 --to=4-17 -- 'X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch\Zielersatz.pdf')""
    $s = "`"$qpdfexe`" '$targetPdfFile'"
    foreach ($r in $replacements) {
        $s += " --overlay '$($r.path)' --from=$($r.pdfPageNumber) --to=$($r.docPageNumber)"
    }
    $s += " -- '$targetPdfFile'"
    $s
    Invoke-Expression $s
}