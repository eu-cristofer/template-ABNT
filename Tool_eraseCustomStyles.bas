Attribute VB_Name = "Tool_eraseCustomStyles"
Sub eraseCustomStyles()

    On Error Resume Next

    ActiveDocument.Styles("New normal").Delete
    ActiveDocument.Styles("Corpo do texto").Delete
    ActiveDocument.Styles("T�tulo n�o numerado").Delete
    ActiveDocument.Styles("T�tulo fora do sum�rio").Delete
    
    For intLoop = 1 To 4
        ActiveDocument.Styles("T�tulo" & " " & intLoop).Delete
    Next intLoop
    
    On Error GoTo 0
    
End Sub
