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

    [boolean] concatenate($path1, $path2, $addPageBreakInBetween) {
        return $this.Word.Run("xToDoc.concatenate", "$path1", "$path2", $addPageBreakInBetween)
    }

    [boolean] concatenate($path1, $path2) {
        return $this.concatenate($path1, $path2, $false)
    }

    [boolean] concatenatePdfPage($path, $pdfFile, $pdfPageNumber, $pdfHeadings) {
        if (-not $this.concatenate($path, $this.templatePdfPage)) { return $false }

        if (-not $this.replaceLastVariable($path, "pdfFile", $pdfFile)) { return $false }
        if (-not $this.replaceLastVariable($path, "pdfPageNumber", $pdfPageNumber)) { return $false }

        $pdfHeadingLines = [System.Collections.ArrayList]@()
        $i = 0
        while ($i -lt $pdfHeadings.Count) {
            if ($pdfHeadings[$i].pdfHeadingTier -ne "None" -or $i -eq $pdfHeadings.Count-1) {
                $pdfHeadingLines.Add("{>pdfHeading$($pdfHeadings[$i].pdfHeadingTier)<}$($pdfHeadings[$i].pdfHeadingText)")
            }
            $i += 1
        }
        if (-not $this.replaceLastVariable($path, "pdfHeadings", $pdfHeadingLines -join "^p")) { return $false }

        return $true
    }

    [boolean] replace($path, $text, $replacement, $replaceAll, $startEnd) {
        return $this.Word.Run("xToDoc.replace", "$path", "$text", "$replacement", $replaceAll, $startEnd)
    }

    [boolean] replaceVariable($path, $name, $value, $replaceAll, $startEnd) {
        $length = 254
        if ($value.Length -gt $length) {
            $variableCount = [Math]::Ceiling($value.Length/$length)
            $restStringLength = $value.Length % $length

            $replacementVariablesString = ""
            for ($i = 0; $i -lt $variableCount; $i++) {
                $replacementVariablesString += "{{`$$($name)_$($i)}}"
            }
            if (-not $this.replaceVariable($path, $name, $replacementVariablesString, $replaceAll, $startEnd)) { return $false }
            if ($replacementVariablesString.Length -gt $length) {
                throw "Der Variableninhalt von `$$name ist viel zu lang oder der Variablenname selbst ist viel zu lang!"
            }

            $oldOffset = 0
            for  ($i = 0; $i -lt $variableCount-1; $i++) {
                $carets = 0
                for ($l = (($i+1)*$length)-1; $l -ge $i*$length+$oldOffset; $l--) {
                    if ($value[$l] -ne '^') { break }
                    $carets += 1
                }

                $offset = $carets % 2

                if (-not $this.replaceVariable($path, "$($name)_$($i)", $value.Substring($i * $length + $oldOffset, $length + ($offset - $oldOffset)), $replaceAll, $startEnd)) { return $false }

                $oldOffset = $offset
            }
            if (-not $this.replaceVariable($path, "$($name)_$($variableCount-1)", $value.Substring(($variableCount-1) * $length + $oldOffset, $restStringLength - $oldOffset), $replaceAll, $startEnd)) { return $false }

            return $true
        }

        return $this.replace("$path", "{{`$$name}}", "$value", $replaceAll, $startEnd)
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

    [boolean] applyFormatting($path) {
        return $this.Word.Run("xToDoc.applyFormatting", "$path")
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