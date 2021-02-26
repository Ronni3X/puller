'Dim the primary objects.

Dim objApp
Dim objIE
Dim objWindow

'Set the primary objects.

Set objApp = CreateObject("Shell.Application")
Set objIE = Nothing
Dim arr_lett
arr_lett = Array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z")



'while loop check
Dim ch
ch = 0

'write to a file
Set objfso = CreateObject("Scripting.FileSystemObject")
outfile = "C:\Users\ACE\Documents\lhd7_2.txt"
Set objfile = objfso.CreateTextFile(outfile,True)

For Each letter in arr_lett
	ch = 0
	For Each objWindow In objApp.Windows
		If (InStr(objWindow.Name, "Internet Explorer")) Then
			Set objIE = objWindow
			Exit For
		End If
	Next
	
	With objIE

			.Visible = True    'Make sure to set the window of IE to visible.

			Do While .Busy Or .readyState <> 4
				'Do nothing, wait for the browser to load.
			Loop

			Do While .Document.ReadyState <> "complete"
				'Do nothing, wait for the VBScript to load the document of the website.
			Loop
			
			'Do stuff
			
			Set fram = .Document.getElementsByTagName("frame")(0)
			Set fram_doc = fram.contentDocument
			
			For Each input In fram_doc.getElementsByTagName("input")
				'WScript.Echo input.Value
				'WScript.Echo input.Name
				If (InStr(input.name,"FORM_PARAM_LNAME")) Then
					input.value = letter
				End If
				
				If (InStr(input.name,"Search")) Then
					'WScript.Echo input.Name
					input.Click
				End If
			Next
			
	End With
	
	While ch=0

		'Identify the IE window and connect.

		For Each objWindow In objApp.Windows
			If (InStr(objWindow.Name, "Internet Explorer")) Then
				Set objIE = objWindow
				Exit For
			End If
		Next

		With objIE

			.Visible = True    'Make sure to set the window of IE to visible.

			Do While .Busy Or .readyState <> 4
				'Do nothing, wait for the browser to load.
			Loop

			Do While .Document.ReadyState <> "complete"
				'Do nothing, wait for the VBScript to load the document of the website.
			Loop
	 
			'***THIS IS WHERE YOU DO SOMETHING***

			Set frame = .Document.getElementsByTagName("frame")(1)
			Set frame_doc = frame.contentDocument
			
			For Each p In frame_doc.getElementsByTagName("p")
				'WScript.Echo p.innerText
				If (InStr(p.innerText,"No entries match the requested search term. Please try a different search.")) Then
					ch=ch+1
				End If
			Next
						
			For Each input In frame_doc.getElementsByTagName("input")
				objfile.Write input.Value & vbCrLf
				'WScript.Echo input.Value
				If (InStr(input.Value,"  Next >>  ")) Then
					'WScript.Echo VarType(input)
					If (input.disabled < 0) Then
						ch=ch+1
					Else
						input.Click
					End If
				End If
			Next
		

		End With 


		'objIE = Nothing 
		'objWindow = Nothing
	Wend

Next

'close the file
objfile.close

'Clear the objects.

On Error Resume Next 

objApp = Nothing 


'archived
  '.Document.getElementsByTitle("Click here to send mail to this address").Item
  'Set href_list = .Document.getElementsByTagName("a")
  'MsgBox .Document.All.Item("a").length

  'WScript.Echo .Document.Body.innerText
  'WScript.Echo frame_doc.getElementsByTagName("a").length
  'For Each link In frame_doc.getElementsByTagName("a")
  '  If (InStr(link,"mailto:")) Then
  '    WScript.Echo Mid(link,8)
  '  End If
  'Next