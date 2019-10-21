Dim oClipboard, sapi, message, oldmessage, i, iterations

i = 0
Set oClipboard = New clsClipboard
message = oClipboard.GetText
oldmessage = message

Set sapi= CreateObject("sapi.spvoice")
Set sapi.Voice = sapi.GetVoices.Item(2)

iterations = CInt(InputBox("How many iterations(1 iteration/s): ","iterations"))
Do Until i=iterations
message = oClipboard.GetText
If message <> oldmessage Then sapi.Speak message
oldmessage = message
i = i+1
WScript.Sleep 1000
Loop

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''  Subs / Classes / Functions  '''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class clsClipboard

''''' This class is used for manupilting the Windows clipboard.
''''' Under some versions of windows you may be prompted with a UAC message box
''''' before continuing.  For an original copy of the example file showing how to
''''' call this class and other interesting scripts goto www.RichSchreiber.com

    Private oHTML
    
    Private Sub Class_Initialize
        Set oHTML = CreateObject("InternetExplorer.Application")
        oHTML.Navigate ("about:blank")
    End Sub


    
    ' Get text from Clipboard   
    Public Property Get GetText()
        GetText = oHTML.Document.ParentWindow.ClipboardData.GetData("Text")
    End Property
    
    
    Private Sub Class_Terminate
        oHTML.Quit
        Set oHTML = Nothing
    End Sub
    
End Class
