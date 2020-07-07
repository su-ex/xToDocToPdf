$pdfMagic = "################################PDF################################"
$pdfFilePrefix = "Datei: "
$pdfPagePrefix = "Seite "

$wdFindStop = 0
$wdReplaceNone = 0
$wdActiveEndPageNumber = 3

$wdExportFormatPDF = 17
$wdExportOptimizeForPrint = 0
$wdExportAllDocument = 0
$wdExportDocumentContent = 0
$wdExportCreateHeadingBookmarks = 1

class WordPdfExportHelper {
    $Word = $Null
    $doc = $Null

    WordPdfExportHelper([string]$sourceWordFile) {
        $this.Word = New-Object -ComObject Word.Application
        $this.Word.Visible = $False
        $this.doc = $this.Word.Documents.Open($sourceWordFile, $False, $True)

        # see: https://www.mrexcel.com/board/threads/selection-find-execute-does-not-work-for-read-only-mode.985285/
        $this.doc.ActiveWindow.View.ReadingLayout = $False
    }

    [boolean] isPortrait() {
        $width = $this.doc.PageSetup.PageWidth
        $height = $this.doc.PageSetup.PageHeight

        Write-Debug "Word document page dimesions: width: $width, height: $height"
        $isPortrait = $width -lt $height

        return $isPortrait
    }

    [System.Collections.ArrayList] getPdfReplacementPages() {
        $pdfReplacements = [System.Collections.ArrayList]@()

        $objSelection = $this.Word.Selection

        $FindText = "$($Script:pdfMagic)^l$($Script:pdfFilePrefix)*^l$($Script:pdfPagePrefix)*^l$($Script:pdfMagic)"
        $MatchCase = $False
        $MatchWholeWord = $False
        $MatchWildCards = $True
        $MatchSoundsLike = $False
        $MatchAllWordForms = $False
        $Forward = $True
        $Wrap = $Script:wdFindStop
        $Format = $False
        $ReplaceWith = ""
        $Replace = $Script:wdReplaceNone

        while ($objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord,
        $MatchWildCards,$MatchSoundsLike,$MatchAllWordForms,$Forward,
        $Wrap,$Format,$ReplaceWith,$Replace)) {
            $objSelection.Text -match "$($Script:pdfMagic)`v$($Script:pdfFilePrefix)(.*)`v$($Script:pdfPagePrefix)(.*)`v$($Script:pdfMagic)" | Out-Null
            $pdfReplacements.Add([PSCustomObject]@{
                path = $matches[1]
                docPageNumber = $objSelection.Information($Script:wdActiveEndPageNumber)
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
            $Script:wdExportFormatPDF,
            $false,
            $Script:wdExportOptimizeForPrint,
            $Script:wdExportAllDocument,
            0,
            0,
            $Script:wdExportDocumentContent,
            $true,
            $true,
            $Script:wdExportCreateHeadingBookmarks,
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