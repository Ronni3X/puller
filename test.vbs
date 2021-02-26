'Dim the primary objects.

Dim objIE

'Set the primary objects.

Set objIE = CreateObject("InternetExplorer.Application")

With objIE

  .Visible = True    'Make sure to set the window of IE to visible.

  .Navigate("https://www.google.com/")    'Navigate to the desired website.

  Do While .Busy Or .readyState <> 4
    'Do nothing, wait for the browser to load.
  Loop

  Do While .Document.ReadyState <> "complete"
    'Do nothing, wait for the VBScript to load the document of the website.
  Loop
 
  '***THIS IS WHERE YOU DO SOMETHING***

End With 

'Clear the objects.

On Error Resume Next

objIE = Nothing