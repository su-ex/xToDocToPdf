$qpdfexe = (Get-ChildItem -Path "$PSScriptRoot\..\..\qpdf-**\" -Recurse | Where-Object { $_.Name -like "qpdf.exe" } | Sort-Object)[-1]
$gsexe = (Get-ChildItem -Path "C:\Program Files\gs\" -Recurse | Where-Object { $_.Name -like "gs*c.exe" } | Sort-Object)[-1]
$gspdfinfo = (Get-ChildItem -Path "C:\Program Files\gs\" -Recurse | Where-Object { $_.Name -like "pdf_info.ps" } | Sort-Object)[-1]

[regex]$pdfSizeExtractionPattern = '(?i)^(?=.*Page\s+(?<page>\d+))(?=.*MediaBox:\s+\[\d+\s+\d+\s+(?<width>\d+(\.\d+)?)\s+(?<height>\d+(\.\d+)?)\])(?=.*Rotate\s+=\s+(?<rotate>\d+))?.*$'

Function getPdfPageNumber($file) {
    return [int](& "$qpdfexe" --show-npages "`"$file`"")
}

Function getPdfPageDimensions($file) {
    [string[]]$lines = (& "$gsexe" -dQUIET -dNODISPLAY -dNOSAFER -q -sFile="$file" "$gspdfinfo")

    $pages = [System.Collections.ArrayList]@()

    foreach ($line in $lines) {
        $m = $pdfSizeExtractionPattern.match($line)
        [int]$page, [int]$width, [int]$height, [int]$rotate = $m.Groups['page', 'width', 'height', 'rotate'].Value

        # pattern not recognized
        if (-not $m.Success) {
            continue
        }

        $isPortrait = $width -lt $height
        if ([Math]::Abs(($rotate / 90) % 2) -eq 1) {
            $isPortrait = !$isPortrait;
        }

        $pages.Add([PSCustomObject]@{
            pageNumber = $page
            width = $width
            height = $height
            rotate = $rotate
            isPortrait = $isPortrait
        }) | Out-Null
    }

    return ($pages | Sort-Object -Property pageNumber)
}

Function overlayPdfPages($targetPdfFile, $replacementGroups) {
    foreach ($rg in $replacementGroups) {
        # input same as output (target pdf file which is overlaid to), group name as pdf file (where the overlays are from), to and from comma seperated page numbers
        (& "$qpdfexe" --replace-input "$targetPdfFile" --overlay "$($rg.Name)" --to=$($rg.Group.docPageNumber -join ',') --from=$($rg.Group.pdfPageNumber -join ',') --)
    }
}