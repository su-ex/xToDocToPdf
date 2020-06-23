$qpdfexe = (Resolve-Path "$PSScriptRoot\..\..\qpdf-**\**\bin\qpdf.exe")

Function getPdfPageNumber($file) {
    #Write-Host "exe: $qpdfexe, file: $file"
    return [int](& "$qpdfexe" --show-npages "`"$file`"")
}