$qpdfexe = (Resolve-Path "$PSScriptRoot\..\..\qpdf-**\**\bin\qpdf.exe")

Function getPdfPageNumber($file) {
    #Write-Host "exe: $qpdfexe, file: $file"
    return [int](& "$qpdfexe" --show-npages "`"$file`"")
}

Function overlayPdfPages($targetPdfFile, $replacements) {
    # one group for each pdf file (more efficient --> faster) 
    $replacementGroups = $replacements | Group-Object -Property path
    
    foreach ($rg in $replacementGroups) {
        # input same as output (target pdf file which is overlaid to), group name as pdf file (where the overlays are from), to and from comma seperated page numbers
        (& "$qpdfexe" --replace-input "$targetPdfFile" --overlay "$($rg.Name)" --to=$($rg.Group.docPageNumber -join ',') --from=$($rg.Group.pdfPageNumber -join ',') --)
    }
}