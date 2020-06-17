Add-Type -AssemblyName System.Windows.Forms

function showTree($description) {
    $ButtonOK_Click = {
        foreach ($d in $description) {
            $d.enabled = $d.asset.checked
            $d.asset = $Null
        }
    }

    $window = New-Object System.Windows.Forms.Form
    $window.ClientSize = '342, 502'
    $window.FormBorderStyle = 'FixedDialog'
    $window.StartPosition = "CenterScreen"
    $window.Text = "Auswahl treffen"

    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Dock = 'Fill'
    $treeView.CheckBoxes = $true

    $treeView.Add_AfterCheck({
        if ($_.Node.Checked) {
            if ($_.Node.Parent) {
                $_.Node.Parent.Checked = $True
            }
        } else {
            if ($_.Node.Nodes.Count -gt 0) {
                foreach ($n in $_.Node.Nodes) {
                    $n.Checked = $False
                }
            }
        }
    })

    [System.Collections.Stack]$stack = @()
    $stack = New-Object System.Collections.Stack
    $stack.Push($treeView)
    $stack.Push($Null)
    $lastIndent = 0
    foreach ($d in $description) {
        $indentDifference = $d.indent - $lastIndent
        #Write-Host "c: $($d.indent), l: $lastIndent, d: $indentDifference"
        for ($i = $indentDifference; $i -le 0; $i++) {
            $stack.Pop() | Out-Null
        }

        $newNode = New-Object System.Windows.Forms.TreeNode
        $newNode.Text = $d.desc
        $newNode.Checked = $d.enabled
        $stack.Peek().Nodes.Add($newNode) | Out-Null
        $d.asset = $newNode

        $stack.Push($d.asset)
        $lastIndent = $d.indent
    }

    foreach ($n in $treeView.Nodes) {
        if ($n.Parent -and !$n.Parent.Checked) {
            $n.Checked = $False
        }
    }

    $treeView.ExpandAll()

    $ButtonOK = New-Object System.Windows.Forms.Button
    $ButtonOK.DialogResult = 'OK'
    $ButtonOK.Location = '245,467'
    $ButtonOK.Size = '75,23'
    $ButtonOK.Name = 'ButtonOK'
    $ButtonOK.Text = 'OK'
    $ButtonOK.add_Click($ButtonOK_Click)
    $window.Controls.Add($ButtonOK)

    $window.Controls.Add($treeView)
    $window.ShowDialog()
}