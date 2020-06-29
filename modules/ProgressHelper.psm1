Import-Module "$PSScriptRoot\HelperFunctions.psm1" -Force

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

class ProgressHelper {
    [String] $activityName
    [String] $currentOperation
    [int] $totalOperations = 100
    [int] $i = 1
    [bool] $wasSuccessful = $false

    # see: https://gallery.technet.microsoft.com/scriptcenter/Progress-Bar-With-d3924344
    [Object] $ObjForm
    [Object] $ObjLabel
    [Object] $PB

    ProgressHelper([String] $activityName) {
        $this.activityName = $activityName
        $this.currentOperation = "Start"

        ## -- Create The Progress-Bar
        $this.ObjForm = New-Object System.Windows.Forms.Form
        $this.ObjForm.Text = $this.activityName
        $this.ObjForm.Height = 100
        $this.ObjForm.Width = 500
        $this.ObjForm.BackColor = "White"

        $this.ObjForm.FormBorderStyle = 'FixedSingle'
        $this.ObjForm.StartPosition = 'CenterScreen'

        ## -- Create The Label
        $this.ObjLabel = New-Object System.Windows.Forms.Label
        $this.ObjLabel.Text = $this.currentOperation
        $this.ObjLabel.Left = 5
        $this.ObjLabel.Top = 10
        $this.ObjLabel.Width = 500 - 20
        $this.ObjLabel.Height = 15
        ## -- Add the label to the Form
        $this.ObjForm.Controls.Add($this.ObjLabel)

        $this.PB = New-Object System.Windows.Forms.ProgressBar
        $this.PB.Name = "PowerShellProgressBar"
        $this.PB.Value = 0
        $this.PB.Style="Continuous"

        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = 500 - 40
        $System_Drawing_Size.Height = 20
        $this.PB.Size = $System_Drawing_Size
        $this.PB.Left = 5
        $this.PB.Top = 40
        $this.ObjForm.Controls.Add($this.PB)

        ## -- Show the Progress-Bar and Start The PowerShell Script
        $this.ObjForm.Show() | Out-Null
        $this.ObjForm.Focus() | Out-NUll
        $this.ObjLabel.Text = $this.currentOperation
        $this.ObjForm.Refresh()
    }

    [void] setTotalOperations([int] $totalOperations) {
        $this.totalOperations = $totalOperations
    }

    [void] reset() {
        $this.i = 1
    }

    [void] update([String] $operation) {
            $this.currentOperation = $operation
            while ($this.totalOperations -lt $this.i -or $this.totalOperations -le 0) { $this.totalOperations += 1 }

            $this.PB.Value = [int](($this.i / $this.totalOperations) * 100)
            $this.ObjLabel.Text = "Aufgabe $($this.i) von $($this.totalOperations): $($this.currentOperation)"
            $this.ObjForm.Refresh()
            
            $this.i++
    }

    [void] error() {
        $this.wasError = $true
        throw "Aufgabe `"$($this.currentOperation)`" konnte nicht ausgeführt werden."
    }

    [void] success() {
        $this.wasSuccessful = $true
    }

    [void] finish() {
        if ($this.wasSuccessful) {
            $this.PB.Value = 100
            $this.ObjLabel.Text = "Abgeschlossen"
            $this.ObjForm.Refresh()
    
            infoBox "Erfolgreich! :-)"
        }

        $this.ObjForm.Close()
    }
}

Function ProgressHelper([String] $activityName) {
    return [ProgressHelper]::new($activityName)
}