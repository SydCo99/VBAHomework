Attribute VB_Name = "Module2"
Sub conditionalformat()
    For x = 2 To 91
        If Cells(x, 11).Value >= 0 Then
            Cells(x, 11).Interior.ColorIndex = 4
        ElseIf Cells(x, 11).Value < 0 Then
            Cells(x, 11).Interior.ColorIndex = 3
        End If
    Next x
    
End Sub
