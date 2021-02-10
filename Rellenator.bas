Sub Rellenator()
Dim tTable As Table
    Dim cCell As Cell
    Dim sTemp As String
    
    sTemp = InputBox("Frase para relleno", "Introduzca la frase a usar")

    For Each tTable In ActiveDocument.Range.Tables
        For Each cCell In tTable.Range.Cells
            If Len(cCell.Range.Text) < 3 Then
                cCell.Range = sTemp
            End If
        Next
    Next
    Set oCell = Nothing
    Set tTable = Nothing
    
End Sub