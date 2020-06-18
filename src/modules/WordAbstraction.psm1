class WordAbstraction {
    [String] $docmPath = "$PSScriptRoot\Makros.docm"

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
        Write-Host "path1: $path1, path2: $path2"
        return $this.Word.Run("xToDoc.concatenate", [ref]"$path1", [ref]"$path2")
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