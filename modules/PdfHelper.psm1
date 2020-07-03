$qpdfexe = ((Get-ChildItem -Path "$PSScriptRoot\..\..\qpdf-**\" -Recurse | Where-Object { $_.Name -like "qpdf.exe" } | Sort-Object -Descending)[0]).FullName

Function getPdfPageNumber($file) {
    $output = (& "$qpdfexe" --show-npages "$file" 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "Calling qpdf failed:`n$output"
    }

    return [int]$output
}

Function getPdfPageDimensions($file) {
    $output = (& "$qpdfexe" "$file" --json 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "Calling qpdf failed:`n$output"
    }

    $pdfjson = $output | ConvertFrom-Json

    $pages = [System.Collections.ArrayList]@()

    for ($i = 0; $i -lt $pdfjson.pages.Count; $i++) {
        $info = $pdfjson.objects.$($pdfjson.pages[$i].object)

        [int]$width = $info.'/MediaBox'[2]
        [int]$height = $info.'/MediaBox'[3]

        [int]$rotate = $info.'/Rotate'

        $isPortrait = $width -lt $height
        if ([Math]::Abs(($rotate / 90) % 2) -eq 1) {
            $isPortrait = !$isPortrait;
        }

        $pages.Add([PSCustomObject]@{
            pageNumber = $i+1
            width = $width
            height = $height
            rotate = $rotate
            isPortrait = $isPortrait
        }) | Out-Null
    }

    return ($pages | Sort-Object -Property pageNumber)
}

Function rotatePdfPages90Deg($file, $pagesToRotate) {
    if ($pagesToRotate.Count -eq 0) { return }
    $output = (& "$qpdfexe" --replace-input "$file" --rotate=+90:$($pagesToRotate -join ',') 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "Calling qpdf failed:`n$output"
    }
}

Function overlayPdfPages($targetPdfFile, $replacementGroups) {
    foreach ($rg in $replacementGroups) {
        # input same as output (target pdf file which is overlaid to), group name as pdf file (where the overlays are from), to and from comma seperated page numbers
        $output = (& "$qpdfexe" --replace-input "$targetPdfFile" --overlay "$($rg.Name)" --to=$($rg.Group.docPageNumber -join ',') --from=$($rg.Group.pdfPageNumber -join ',') -- 2>&1)
        if ($LASTEXITCODE -ne 0) {
            throw "Calling qpdf failed:`n$output"
        }
    }
}