Function getDescription($path) {
    $pieces = [System.Collections.ArrayList]@()

    $lastIndent = 0;
    foreach($line in [System.IO.File]::ReadLines($path)) {
        $desc, $path = $line -split ": "

        $disabled = $False
        if ($desc.SubString(0, 1) -eq ";") {
            $disabled = $True
            $desc = $desc.SubString(1, $desc.length-1)
        }

        $indent = 0
        while ($desc.SubString(0, 1) -eq "`t") {
            $indent++
            $desc = $desc.SubString(1, $desc.length-1)
        }

        if ($indent - $lastIndent -gt 1) {
            throw "Malformed description (indentation mistake)!"
        }

        $pieces.Add([PSCustomObject]@{
            desc = $desc
            path = $path
            enabled = !$disabled
            indent = $indent
            asset = $Null
        }) | Out-Null
    }

    return $pieces
}