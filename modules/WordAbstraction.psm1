class WordAbstraction {
    [String] $docmPath = "$PSScriptRoot\Makros.docm"
    [String] $templatePdfPage = "$PSScriptRoot\..\assets\PDF.docx"

    $Word = $Null
    $docm = $Null

    WordAbstraction() {
        $this.Word = New-Object -ComObject Word.Application
        $this.Word.Visible = $False
        $this.docm = $this.Word.Documents.Open($this.docmPath, $False, $True)
    }

    [boolean] concatenate($path1, $path2) {
        return $this.Word.Run("xToDoc.concatenate", "$path1", "$path2")
    }

    [boolean] concatenatePdfPage($path, $pdfFile, $pdfPageNumber, $pdfHeadingTier, $pdfHeadingText) {
        if (-not $this.concatenate($path, $this.templatePdfPage)) { return $false }
        if (-not $this.replaceLastVariable($path, "pdfFile", $pdfFile)) { return $false }
        if (-not $this.replaceLastVariable($path, "pdfPageNumber", $pdfPageNumber)) { return $false }
        if (-not $this.replaceLastVariable($path, "pdfHeadingTier", $pdfHeadingTier)) { return $false }
        if (-not $this.replaceLastVariable($path, "pdfHeadingText", $pdfHeadingText)) { return $false }
        return $true
    }

    [boolean] replace($path, $text, $replacement, $replaceAll, $startEnd) {
        return $this.Word.Run("xToDoc.replace", "$path", "$text", "$replacement", $replaceAll, $startEnd)
    }

    [boolean] replaceVariable($path, $name, $value, $replaceAll, $startEnd) {
        return $this.Word.Run("xToDoc.replace", "$path", "{{`$$name}}", "$value", $replaceAll, $startEnd)
    }

    [boolean] replaceVariable($path, $name, $value) {
        return $this.replaceVariable("$path", $name, $value, $true, $false)
    }

    [boolean] replaceLastVariable($path, $name, $value) {
        return $this.replaceVariable("$path", $name, $value, $false, $true)
    }

    [boolean] updateHeadings($path) {
        return $this.Word.Run("xToDoc.updateHeadings", "$path")
    }

    [boolean] updateFields($path) {
        return $this.Word.Run("xToDoc.updateFields", "$path")
    }

    [boolean] saveAndClose($path) {
        return $this.Word.Run("xToDoc.saveAndClose", "$path")
    }

    [void] destroy() {
        $this.docm.close($False)
        $this.Word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Word) | Out-Null
    }
}

Function WordAbstraction() {
    return [WordAbstraction]::new()
}