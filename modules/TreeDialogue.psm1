Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# see: https://stackoverflow.com/questions/27904837/powershell-net-collect-all-checked-nodes
function showTree($description) {
    $ButtonOK_Click = {
        foreach ($d in $description) {
            $d.enabled = $d.asset.checked
            $d.asset = $Null
        }
    }

    Function disableChildsOfUnckeckedParents($node, $recurse = $true) {
        foreach ($n in $node.Nodes) {
            if ($n.Parent -and !$n.Parent.Checked) {
                $n.Checked = $False
            }
            if ($recurse) { disableChildsOfUnckeckedParents $n $recurse }
        }
    }

    $window = New-Object System.Windows.Forms.Form
    $window.ClientSize = [System.Drawing.Size]::new(342, 502)
    $window.FormBorderStyle = 'FixedDialog'
    $window.StartPosition = "CenterScreen"
    $window.Text = "Auswahl treffen"

    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Dock = 'Fill'
    $treeView.CheckBoxes = $true

    $treeView.Add_AfterCheck({
        if ($_.Node.Checked) { # enable parents too
            if ($_.Node.Parent) {
                $_.Node.Parent.Checked = $True
            }
        } else { # disable childs too
            # will recurse by itself
            disableChildsOfUnckeckedParents $_.Node $false
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

    # there might be childs of disabled parents, disable the childs too before showing tree
    disableChildsOfUnckeckedParents $treeView.Nodes

    $treeView.ExpandAll()

    $ButtonOK = New-Object System.Windows.Forms.Button
    $ButtonOK.DialogResult = 'OK'
    $ButtonOK.Location = [System.Drawing.Point]::new(245, 467)
    $ButtonOK.Size = [System.Drawing.Size]::new(75, 23)
    $ButtonOK.Name = 'ButtonOK'
    $ButtonOK.Text = 'OK'
    $ButtonOK.add_Click($ButtonOK_Click)
    $window.Controls.Add($ButtonOK)

    $window.Controls.Add($treeView)
    $window.ShowDialog()

    return ("$($window.DialogResult)" -eq "OK")
}