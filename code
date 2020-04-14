Private Sub worksheet_Change(ByVal Target As Range)

Dim col As Range

    With ThisWorkbook.Sheets("sheet1")
    
        .Unprotect "1234"
        
        For Each col In .UsedRange.Columns
        
            col.EntireColumn.Locked = col.Range("A1").Value <> Date
        
        Next col
        
        .Protect "1234"
    
        .EnableSelection = x1NoRestrictions
    End With
    
            

End Sub
