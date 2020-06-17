$Word = New-Object -ComObject Word.Application
$Word.Visible = $False
$myPath = "C:\Users\Lager\Desktop\TestBedienhandbuch"
$targetPath = "$myPath\Ziel.docx"

$newDoc = $Word.Documents.Open("$myPath\FormatvorlagenUndAnfang.docx")
$tempDoc = $Word.Documents.Open("$myPath\WichtigeInformation.doc")
$tempDoc.Range.Copy
$rng = $newDoc.Range
$rng.Collapse($Word.WdCollapseDirection.wdCollapseEnd)
$rng.Paste
$tempDoc.Close($False)

$pathList = New-Object "System.Collections.ArrayList"
$pathList.Add("FormatvorlagenUndAnfang.docx")
$pathList.Add("WichtigeInformation.doc")
#pathList.Add "blablabla.doc"
$pathList.Add("Rest.doc")

<#
$i = 0
ForEach ($element In $pathList) {
Write-Host "count: $i, path: $myPath\$element"
	If ($i -eq 0) {
		$Script:newDoc = $Word.Documents.Open("$myPath\$element")
	} Else {
		$tempDoc = $Word.Documents.Open("$myPath\$element")
		$tempDoc.Content.Copy
		$rng = $newDoc.Content
		$rng.Collapse($Word.WdCollapseDirection.wdCollapseEnd)
		$rng.Paste
		$tempDoc.Close($False)
	}
	$i++
}
    Dim oStory As Range
    For Each oStory In newDoc.StoryRanges
        oStory.Fields.Update
        If oStory.StoryType <> wdMainTextStory Then
            While Not (oStory.NextStoryRange Is Nothing)
                Set oStory = oStory.NextStoryRange
                oStory.Fields.Update
            Wend
        End If
    Next oStory
    Set oStory = Nothing

    Dim TOC As TableOfContents
    For Each TOC In newDoc.TablesOfContents
        TOC.Update
    Next
    
    
    
'    newDoc.UpdateStyles
#>

$newDoc.SaveAs($targetPath)
$newDoc.Close($False)
