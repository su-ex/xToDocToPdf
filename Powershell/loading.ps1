function Call-dialogue_psf {
    Add-Type -AssemblyName System.Windows.Forms

    $window = New-Object System.Windows.Forms.Form
    $window.Size = New-Object Drawing.Size @(400,75)
    $window.StartPosition = "CenterScreen"
    $window.Text = "Generiere Word-Dokument ..."

    $ProgressBar1 = New-Object System.Windows.Forms.ProgressBar
    $ProgressBar1.Location = New-Object System.Drawing.Point(10, 10)
    $ProgressBar1.Size = New-Object System.Drawing.Size(365, 20)
    $ProgressBar1.Style = "Marquee"
    $ProgressBar1.MarqueeAnimationSpeed = 24
    $window.Controls.Add($ProgressBar1)

    Start-Job -ScriptBlock { $window.ShowDialog() }
}
Call-dialogue_psf | Out-Null
Write-Host "blabla"