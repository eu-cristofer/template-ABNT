Attribute VB_Name = "a1_createAbntStyles"
Sub createAbntStyles(control As IRibbonControl)

    Call eraseCustomStyles
 
    Call reducePriority
    
    Call createFontSizeName(styleName:="New normal")
    
    Call createBody(styleName:="Corpo do texto")
    
    Call createNonNumberedTitle(styleName:="Título fora do sumário")
    
    Call createNonNumberedTitle(styleName:="Título não numerado")
    
    Call createNumberedTitle(styleName:="Título")
    
    Call create_list
  
End Sub

Sub reducePriority()

    Set oStyles = Application.ActiveDocument.Styles
    
    For Each Style In oStyles
        If Style.Priority < 50 Then
            Style.Priority = (Style.Priority + 50)
        End If
    Next Style

End Sub

Sub createFontSizeName( _
    styleName As String, _
    Optional ByRef fontSize As Integer = 12, _
    Optional ByRef fontName As String = "Arial" _
    )
'
' Cascade level 0
'
    Set oStyles = _
        Application.ActiveDocument.Styles

    oStyles.Add _
        Name:=styleName, Type:=wdStyleTypeParagraph
       
    With oStyles(styleName)
        .BaseStyle = "No Spacing"
        .QuickStyle = True
        .UnhideWhenUsed = False
        .Visibility = False
        .Priority = 1
        .LanguageID = 1046
        AutomaticallyUpdate = True
        
        With .Font
            .Name = fontName
            .Size = fontSize
        End With
        
    End With

End Sub

Sub createBody( _
    styleName As String, _
    Optional ByRef fontSize As Integer = 12, _
    Optional ByRef fontName As String = "Arial" _
    )
'
' Cascade level 1
'
    Call createFontSizeName(styleName, fontSize, fontName)

    With Application.ActiveDocument.Styles(styleName)
        
        .NextParagraphStyle = "Corpo do texto"
        
        With .ParagraphFormat
        
            .Alignment = wdAlignParagraphJustify
            .LineSpacing = LinesToPoints(1.5)
            .SpaceAfter = LinesToPoints(1)
                
        End With
        
    End With

End Sub

Sub createNonNumberedTitle( _
    styleName As String, _
    Optional ByRef fontSize As Integer = 12, _
    Optional ByRef fontName As String = "Arial" _
    )
'
' Cascade level 2
'
    Call createBody(styleName, fontSize, fontName)

    With Application.ActiveDocument.Styles(styleName)
        
        .Priority = 2
        
        With .Font
            .Bold = True
            .AllCaps = True
        End With
        
        With .ParagraphFormat
             .Alignment = wdAlignParagraphLeft
             .SpaceBefore = LinesToPoints(2)
             .SpaceAfter = LinesToPoints(2)
        End With
        
    End With

End Sub

Sub createNumberedTitle( _
    styleName As String, _
    Optional ByRef fontSize As Integer = 12, _
    Optional ByRef fontName As String = "Arial" _
    )
'
' Cascade level 3
'
    '
    ' Create a new list template
    '
    
    Dim strFormat As String
    strFormat = ""
    
    For intLoop = 1 To 4
        

        
        Call createNonNumberedTitle(styleName & " " & intLoop, fontSize, fontName)
        
        With ActiveDocument.Styles(styleName & " " & intLoop)
            .Priority = 3
 
                
            Select Case intLoop
            
                Case 1
                    With .ParagraphFormat
                        .SpaceBefore = LinesToPoints(2)
                        .SpaceAfter = LinesToPoints(2)
                    End With
                
                Case 2
                    With .ParagraphFormat
                        .SpaceBefore = LinesToPoints(2)
                        .SpaceAfter = LinesToPoints(1)
                    End With
                                
                Case 3
                    With .ParagraphFormat
                        .SpaceBefore = LinesToPoints(1.5)
                        .SpaceAfter = LinesToPoints(0)
                    End With
                    
                    With .Font
                        .Bold = False
                        .AllCaps = True
                    End With
                                
                Case 4
                    With .ParagraphFormat
                        .SpaceBefore = LinesToPoints(1.5)
                        .SpaceAfter = LinesToPoints(0)
                    End With
                    
                    With .Font
                        .Bold = False
                        .AllCaps = False
                    End With
                                
            End Select
        End With
    Next intLoop
    
End Sub
