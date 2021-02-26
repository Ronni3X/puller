Dim objApp
Dim objIE
Dim objWindow

'Set the primary objects.

Set objApp = CreateObject("Shell.Application")
Set objIE = Nothing

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
		
		Set frame = .Document.getElementsByTagName("frame")(1)
		Set frame_doc = frame.contentDocument
		
		For Each p In frame_doc.getElementsByTagName("p")
			WScript.Echo p.innerText
		Next
		
		Set fram = .Document.getElementsByTagName("frame")(0)
		Set fram_doc = fram.contentDocument
		
		For Each opt In fram_doc.getElementsByTagName("option")
			'WScript.Echo opt.value
		Next
		
		For Each input In fram_doc.getElementsByTagName("input")
				'WScript.Echo input.Value
				'WScript.Echo input.Name
				If (InStr(input.name,"Search")) Then
					'WScript.Echo input.Name
					'If (input.disabled < 0) Then
					'	ch=ch+1
					'Else
					'	input.Click
					'End If
				End If
				
				If (InStr(input.name,"FORM_PARAM_LNAME")) Then
					WScript.Echo input.value
					input.value = "C"
					WScript.Echo input.value
				End If
		Next
		
End With