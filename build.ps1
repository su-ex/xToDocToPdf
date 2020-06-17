ls "$PSScriptRoot\*.ps1" | %{
    Invoke-ps2exe "$($_.Fullname)" "$($_.Fullname -replace '.ps1','.exe')" -verbose -x86 -noConsole -noOutput
}