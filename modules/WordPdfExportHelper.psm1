$pdfMagic = "################################PDF################################"
$pdfFilePrefix = "Datei: "
$pdfPagePrefix = "Seite "

# see: https://stackoverflow.com/questions/53170039/powershell-unable-to-find-type-microsoft-office-interop-word-wdsaveformat
$WdTypes = Add-Type -AssemblyName 'Microsoft.Office.Interop.Word' -PassThru
$WdExportFormat = $wdTypes | Where-Object {$_.Name -eq "WdExportFormat"}
$WdExportOptimizeFor = $wdTypes | Where-Object {$_.Name -eq "WdExportOptimizeFor"}
$WdExportRange = $wdTypes | Where-Object {$_.Name -eq "WdExportRange"}
$WdExportItem = $wdTypes | Where-Object {$_.Name -eq "WdExportItem"}
$WdExportCreateBookmarks = $wdTypes | Where-Object {$_.Name -eq "WdExportCreateBookmarks"}

class WordPdfExportHelper {
    $Word = $Null
    $doc = $Null

    WordPdfExportHelper([string]$sourceWordFile) {
        $this.Word = New-Object -ComObject Word.Application
        $this.Word.Visible = $False
        $this.doc = $this.Word.Documents.Open($sourceWordFile, $False, $True)
    }

    [System.Collections.ArrayList] getPdfReplacementPages() {
        $pdfReplacements = [System.Collections.ArrayList]@()

        $objSelection = $this.Word.Selection

        $wdFindStop = 0
        $wdReplaceNone = 0
        $wdActiveEndPageNumber = 3

        $FindText = "$($Script:pdfMagic)^l$($Script:pdfFilePrefix)*^l$($Script:pdfPagePrefix)*^l$($Script:pdfMagic)"
        $MatchCase = $False
        $MatchWholeWord = $False
        $MatchWildCards = $True
        $MatchSoundsLike = $False
        $MatchAllWordForms = $False
        $Forward = $True
        $Wrap = $wdFindStop
        $Format = $False
        $ReplaceWith = ""
        $Replace = $wdReplaceNone

        while ($objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord,
        $MatchWildCards,$MatchSoundsLike,$MatchAllWordForms,$Forward,
        $Wrap,$Format,$ReplaceWith,$Replace)) {
            $objSelection.Text -match "$($Script:pdfMagic)`v$($Script:pdfFilePrefix)(.*)`v$($Script:pdfPagePrefix)(.*)`v$($Script:pdfMagic)" | Out-Null
            $pdfReplacements.Add([PSCustomObject]@{
                path = $matches[1]
                docPageNumber = $objSelection.Information($wdActiveEndPageNumber)
                pdfPageNumber = $matches[2]
            }) | Out-Null
        }

        return $pdfReplacements
    }

    [void] hidePlaceholders() {
        foreach ($shp in $this.doc.Shapes) {
            if ($shp.Title -eq "{>PdfPlaceholder<}") {
                $shp.Fill.Transparency = 0
            }
        }
    }

    [void] export($targetPdfFile) {
        # taken from recorded Word macro and https://stackoverflow.com/questions/57502233/how-to-set-parameters-for-saveas-dialog-in-word-application
        $this.doc.ExportAsFixedFormat(
            $targetPdfFile,
            $Script:WdExportFormat::wdExportFormatPDF,
            $false,
            $Script:WdExportOptimizeFor::wdExportOptimizeForPrint,
            $Script:WdExportRange::wdExportAllDocument,
            0,
            0,
            $Script:WdExportItem::wdExportDocumentContent,
            $true,
            $true,
            $Script:WdExportCreateBookmarks::wdExportCreateHeadingBookmarks,
            $true,
            $true,
            $false
        )
    }

    [void] destroy() {
        $this.doc.close($False)
        $this.Word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Word) | Out-Null
    }
}

Function WordPdfExportHelper($sourceWordFile) {
    return [WordPdfExportHelper]::new($sourceWordFile)
}