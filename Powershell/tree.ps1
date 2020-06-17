Add-Type -AssemblyName System.Windows.Forms

function GetCheckedNode {
    param($nodes)

    foreach ($n in $Nodes) {
        if ($n.checked) {
            Write-Host $n.Text
        }
        if ($n.nodes.count -gt 0)
        {
            GetCheckedNode $n.nodes
        }
    }   
}
$ButtonOK_Click = {
    GetCheckedNode $treeView.Nodes
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

$N1 = $treeView.Nodes.Add('Node 1')
$N2 = $treeView.Nodes.Add('Node 2')
$N3 = $treeView.Nodes.Add('Node 3')

$N1.Checked = $True

$newNode = New-Object System.Windows.Forms.TreeNode
$newNode.Text = 'Node 1 Sub 1'
$newNode.Checked = $True
$N1.Nodes.Add($newNode) | Out-Null

$newNode = New-Object System.Windows.Forms.TreeNode
$newNode.Text = 'Node 1 Sub 2'
$N1.Nodes.Add($newNode) | Out-Null

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