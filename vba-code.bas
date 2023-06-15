Private Sub worksheet_change(ByVal target As Range)
        If target.Column = 3 Then
        Application.EnableEvents = False
        Cells(target.Row, 4).Value = Date
        Application.EnableEvents = True
        
    End If
End Sub
