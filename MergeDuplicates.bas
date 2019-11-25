Attribute VB_Name = "MergeDuplicates"
Sub Merge_Duplicates()

Dim lastrow As Long

' Find the last row with content
lastrow = Cells(Rows.Count, "A").End(xlUp).Row

' Loop through rows backward and
For x = lastrow To 2 Step -1
    For y = 2 To lastrow
        
        ' Merge all cells with same ID, not case sensitive
        If LCase(Cells(x, 1).Value) = LCase(Cells(y, 1).Value) And x > y Then
            Cells(y, 2).Value = Cells(x, 2).Value + Cells(y, 2).Value
            Cells(y, 3).Value = Cells(x, 3).Value + Cells(y, 3).Value
            Cells(y, 4).Value = Cells(x, 4).Value + Cells(y, 4).Value
            Cells(y, 5).Value = Cells(x, 5).Value + Cells(y, 5).Value
            
            ' Delete Duplicate Row
            Rows(x).EntireRow.Delete
            Exit For
        End If
    Next y
Next x

End Sub

