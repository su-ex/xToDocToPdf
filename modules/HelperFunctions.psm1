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

# see: https://mcpmag.com/articles/2016/06/09/display-gui-message-boxes-in-powershell.aspx
Add-Type -AssemblyName PresentationFramework

Function infoBox($message) {
    [System.Windows.MessageBox]::Show($message, 'Information', 'OK', 'Information')
}

Function yesNoBox($title, $message, $default = 'Yes', $icon = 'Question') {
    return [System.Windows.MessageBox]::Show($message, $title, 'YesNo', $icon, $default)
}

Function exitError($message) {
    if ($message) {
        [System.Windows.MessageBox]::Show($message, 'Fehlgeschlagen', 'OK', 'Stop')
    }
    exit -1
}