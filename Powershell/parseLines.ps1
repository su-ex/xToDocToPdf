$pieces = [System.Collections.ArrayList]@()

foreach($line in [System.IO.File]::ReadLines("X:\Vorlagen\Bedienhandbuch\Vorlage.desc")) {
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

    $pieces.Add([PSCustomObject]@{
        desc = $desc
        path = $path
        disabled = $disabled
        indent = $indent
        asset = $Null
    }) | Out-Null
}

$pieces | Format-Table