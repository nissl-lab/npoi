Sub TestMacro()
'
' TestMacro Macro
' This is a test macro
'

'
    With ActivePresentation.Slides(1).Shapes(1)
        .Title.TextFrame.TextRange.Text = "This is a macro slideshow"
    End With
    With ActivePresentation.Slides(1).Shapes(2)
        .Title.TextFrame.TextRange.Text = "subtitle"
    EndWith
End Sub
