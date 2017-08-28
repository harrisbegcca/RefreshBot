On Error Resume Next

Set objExplorer = CreateObject("InternetExplorer.Application")


objExplorer.Navigate "http://www.example.com/"

objExplorer.Visible = 1


Wscript.Sleep 1


Set objDoc = objExplorer.Document


Do While True

    Wscript.Sleep 300

    objDoc.Location.Reload(True)

    If Err <> 0 Then

        Wscript.Quit

    End If

Loop
