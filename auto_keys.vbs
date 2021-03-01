'read from all lines from file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("C:\Users\nrt\Documents\ships\test.txt")
Set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep 5000

'loop through with sleep
'keep 50 millisecond sleep in there so that wscript doesn't mess up
row = 0
Do Until file.AtEndOfStream
	line = file.ReadLine
	WshShell.SendKeys line
	WshShell.SendKeys "{ENTER}"
	WScript.Sleep 50
	row = row + 1
Loop

file.Close