Attribute VB_Name = "a2_createAbntTOC"
Sub createAbntTOC(control As IRibbonControl)

    'Set myRange = ActiveDocument.Range(0, 0)
    Set myRange = Selection.Range
    
    stylesList = "T�tulo n�o numerado,1"
    
    For intLoop = 1 To 4
        stylesList = stylesList + ", T�tulo" & " " & intLoop & "," & intLoop
        
        With ActiveDocument.Styles("TOC " & intLoop)
            .AutomaticallyUpdate = True
            .BaseStyle = "New normal"
            .NextParagraphStyle = "New normal"
            
            Select Case intLoop
            
                Case 1
                    .Font.AllCaps = True
                    .Font.Bold = True
                Case 2
                    .Font.AllCaps = True
                    .Font.Bold = False
                Case 3
                    .Font.AllCaps = False
                    .Font.Bold = False
                Case 4
                    .Font.AllCaps = False
                    .Font.Bold = False
                    .Font.Size = 11
            End Select
            
        End With
    
    Next intLoop
    
    ActiveDocument.TablesOfContents.Add _
        Range:=myRange, _
        UseFields:=False, _
        UseHeadingStyles:=True, _
        LowerHeadingLevel:=4, _
        UpperHeadingLevel:=1, _
        AddedStyles:= _
            stylesList

End Sub
    
