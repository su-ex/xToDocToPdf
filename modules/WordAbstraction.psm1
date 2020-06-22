class WordAbstraction {
    [String] $docmPath = "$PSScriptRoot\Makros.docm"
    [String] $templatePdfPage = "$PSScriptRoot\..\assets\PDF.docx"

    $Word = $Null
    $docm = $Null

    $tempPath1 = ""
    $tempPath2 = ""

    WordAbstraction() {
        $this.Word = New-Object -ComObject Word.Application
        $this.Word.Visible = $False
        $this.docm = $this.Word.Documents.Open($this.docmPath, $False, $True)
    }

    [void] add([String]$operation, [ScriptBlock]$job) {
        $this.jobs.Enqueue([PSCustomObject]@{
            operation = $operation
            job = $job
        })
    }

    [boolean] concatenate($path1, $path2) {
        #Write-Host "path1: $path1, path2: $path2"
        return $this.Word.Run("xToDoc.concatenate", [ref]"$path1", [ref]"$path2")
    }

    [boolean] concatenatePdfPage($path, $pdfFile, $pdfPageNumber) {
        if (-not $this.concatenate($path, $this.templatePdfPage)) { return $false }
        if (-not $this.replaceLastVariable($path, "pdfFile", $pdfFile)) { return $false }
        if (-not $this.replaceLastVariable($path, "pdfPageNumber", $pdfPageNumber)) { return $false }
        return $true
    }

    [boolean] replace($path, $text, $replacement, $replaceAll, $startEnd) {
        return $this.Word.Run("xToDoc.replace", [ref]"$path", [ref]"$text", [ref]"$replacement", [ref]$replaceAll, [ref]$startEnd)
    }

    [boolean] replaceVariable($path, $name, $value, $replaceAll, $startEnd) {
        return $this.Word.Run("xToDoc.replace", [ref]"$path", [ref]"{{`$$name}}", [ref]"$value", [ref]$replaceAll, [ref]$startEnd)
    }

    [boolean] replaceVariable($path, $name, $value) {
        return $this.replaceVariable("$path", $name, $value, $true, $false)
    }

    [boolean] replaceLastVariable($path, $name, $value) {
        return $this.replaceVariable("$path", $name, $value, $false, $true)
    }

    [boolean] updateHeadings($path) {
        return $this.Word.Run("xToDoc.updateHeadings", [ref]"$path")
    }

    [boolean] updateFields($path) {
        return $this.Word.Run("xToDoc.updateFields", [ref]"$path")
    }

    [boolean] saveAndClose($path) {
        return $this.Word.Run("xToDoc.saveAndClose", [ref]"$path")
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