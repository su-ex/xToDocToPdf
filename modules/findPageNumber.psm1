$pdfMagic = "################################PDF################################"
$pdfFilePrefix = "Datei: "
$pdfPagePrefix = "Seite "

$docPath = "X:\Skripte\xToDoc\assets\PDF.docx"

$pdfReplacements = [System.Collections.ArrayList]@()

$objWord = New-Object -comobject Word.Application
$objWord.Visible = $False
$objDoc = $objWord.Documents.Open($docPath, $False, $True)
$objSelection = $objWord.Selection

$wdFindStop = 0
$wdReplaceNone = 0
$wdActiveEndPageNumber = 3

$FindText = "$pdfMagic^l$pdfFilePrefix*^l$pdfPagePrefix*^l$pdfMagic"
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
    $objSelection.Text -match "$pdfMagic`v$pdfFilePrefix(.*)`v$pdfPagePrefix(.*)`v$pdfMagic" | Out-Null
    $pdfReplacements.Add([PSCustomObject]@{
        path = $matches[1]
        docPageNumber = $objSelection.Information($wdActiveEndPageNumber)
        pdfPageNumber = $matches[2]
    }) | Out-Null
}

$objDoc.Close()
$objWord.Quit()

$pdfReplacements