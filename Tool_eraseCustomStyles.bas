Attribute VB_Name = "Tool_eraseCustomStyles"
Sub eraseCustomStyles()

    On Error Resume Next

    ActiveDocument.Styles("New normal").Delete
    ActiveDocument.Styles("Corpo do texto").Delete
    ActiveDocument.Styles("Título não numerado").Delete
    ActiveDocument.Styles("Título fora do sumário").Delete
    
    For intLoop = 1 To 4
        ActiveDocument.Styles("Título" & " " & intLoop).Delete
    Next intLoop
    
    On Error GoTo 0
    
End Sub
