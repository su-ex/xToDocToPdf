# see: https://germanpowershell.com/stack-lifo-verwenden/

## STACK ##            
<#
Einfache LIFO (Last in, first out) Klasse.
Kann zur Datenverarbeitung bzw. als Warteschlange verwendet werden 
Dient zur Automatisierung von Abläufen und Daten.
#>            
            
# Stack erstellen            
[System.Collections.Stack]$MyStack = @()            
$MyStack = New-Object System.Collections.Stack            
            
# Werte zum Stack hinzufügen            
$MyStack.Push("PowerSHELL")            
$MyStack.Push("ist")            
$MyStack.Push("super")            
$MyStack.Push("!")            
            
# Letzter Wert ausgeben            
$MyStack.Peek()            
            
# Grösse ausgeben            
$MyStack.Count            
            
# Letzter Wert ausgeben und entfernen            
$MyStack.Pop()            
            
# Prüfen ob bestimmter Wert vorhanden ist            
$MyStack.Contains("ist") 
            
# Letzter Wert ausgeben und entfernen            
$MyStack.Pop()    
            
# Letzter Wert ausgeben und entfernen            
$MyStack.Pop()    
            
# Letzter Wert ausgeben und entfernen            
$MyStack.Pop()               
            
# Stack komplett leeren            
$MyStack.Clear()
            
# Grösse ausgeben            
$MyStack.Count     
            
# Prüfen ob bestimmter Wert vorhanden ist            
$MyStack.Contains("ist") 

