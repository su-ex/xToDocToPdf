$qpdfexe = (Resolve-Path "$PSScriptRoot\..\..\qpdf-**\**\bin\qpdf.exe")

Function getPdfPageNumber($file) {
    #Write-Host "exe: $qpdfexe, file: $file"
    return [int](& "$qpdfexe" --show-npages "`"$file`"")
}

Function overlayPages($replacements){
    return [int](& "$qpdfexe" 'X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch\Ziel.pdf' --overlay 'X:\Vorlagen\Bedienhandbuch\de\SaudiArab60-100.pdf' --from=1-14 --to=4-17 -- 'X:\Projekte\2020\PR-2000158_IMB Stromversorgungssysteme GmbH_Test Bedienhandbuch\TestBedienhandbuch\Zielersatz.pdf')
}