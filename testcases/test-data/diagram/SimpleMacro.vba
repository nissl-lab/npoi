Sub TestMacro()
'
' TestMacro Macro
' This is a test macro
'

'
    With ActiveDocument.Pages(1).Shapes(1)
        .Text = "This is a macro vector graphics drawing"
    End With
End Sub

