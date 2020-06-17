$Document = "$PSScriptRoot\Makros.docm"

$Text = "2532-00"
$ReplaceText = "2532-35"

$Word = new-object -ComObject Word.Application
$Word.Visible = $False
$doc = $Word.Documents.Open($Document)
$arg1=[ref]"Hello"
$arg2=[ref]", Buddy!"
$Word.Run("PrintAll",$arg1,$arg2) # The returned value is Hello,Buddy
$argx = [ref]@($arg1,$arg2)
$Word.Run("AnyNumberArgs",$argx)
$doc.close()
$Word.Quit()
$a=[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word)