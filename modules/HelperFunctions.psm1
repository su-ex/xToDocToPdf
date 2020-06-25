# see: https://stackoverflow.com/questions/25682507/powershell-inline-if-iif
Function IIf($If, $Right, $Wrong) {If ($If) {$Right} Else {$Wrong}}

Function ItIf($If, $Wrong) {If ($If) {$If} Else {$Wrong}}

# see: https://powershellexplained.com/2017-01-13-powershell-variable-substitution-in-strings/
Function replaceTokens([String]$placeholderString, [hashtable]$tokens) {
    foreach ($token in $tokens.GetEnumerator()) {
        $pattern = '\$\{' + $token.Key + '\}'
        $placeholderString = $placeholderString -replace $pattern, $token.Value
    }
    return $placeholderString
}

Function makePathAbsolute([String]$base, [String]$child) {
    if (-not [System.IO.Path]::IsPathRooted($child)) {
        return [System.IO.Path]::GetFullPath((Join-Path -Path $base -ChildPath $child))
    } else {
        return $child
    }
}