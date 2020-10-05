# see: https://powershellexplained.com/2017-01-13-powershell-variable-substitution-in-strings/
Function replaceTokens([String]$placeholderString, [hashtable]$tokens) {
    foreach ($token in $tokens.GetEnumerator()) {
        $pattern = '\$\{' + $token.Key + '\}'
        $placeholderString = $placeholderString -replace $pattern, $token.Value
    }
    return $placeholderString
}


Function replaceEachInString([string] $s, [array] $replacements) {
    foreach ($replacement in $replacements) {
        $s = $s -replace $replacement
    }
    return $s
}



Function makePathAbsolute([String]$base, [String]$child) {
    if (-not [System.IO.Path]::IsPathRooted($child)) {
        return [System.IO.Path]::GetFullPath((Join-Path -Path $base -ChildPath $child))
    } else {
        return $child
    }
}



# see: https://stackoverflow.com/questions/46276418/how-to-follow-a-shortcut-in-powershell
function Get-ShortcutTargetPath($fileName) {
    $sh = New-Object -COM WScript.Shell
    $targetPath = $sh.CreateShortcut($fileName).TargetPath 
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sh) | Out-Null
    return $targetPath
}

# see: https://stackoverflow.com/a/24992975
function Test-FileLock {
    param (
      [parameter(Mandatory=$true)][string]$Path
    )
  
    $oFile = New-Object System.IO.FileInfo $Path
  
    if ((Test-Path -Path $Path) -eq $false) {
      return $false
    }
  
    try {
      $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
  
      if ($oStream) {
        $oStream.Close()
      }
      return $false
    } catch {
      # file is locked by a process.
      return $true
    }
  }


### dialogue boxes

# see: https://mcpmag.com/articles/2016/06/09/display-gui-message-boxes-in-powershell.aspx
Add-Type -AssemblyName PresentationFramework

Function infoBox($message) {
    [System.Windows.MessageBox]::Show($message, 'Information', 'OK', 'Information') | Out-Null
}

Function yesNoBox($title, $message, $default = 'Yes', $icon = 'Question') {
    return [System.Windows.MessageBox]::Show($message, $title, 'YesNo', $icon, $default)
}

Function exitError($message) {
    if ($message) {
        [System.Windows.MessageBox]::Show($message, 'Fehlgeschlagen', 'OK', 'Stop') | Out-Null
    }
    exit -1
}