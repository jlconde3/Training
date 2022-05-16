Sub PLANNING_MP()

Dim i, colum As Integer
Dim day As Date
Dim k As Long

With Worksheets("PLANNING_MP")
k = .Range("A" & Rows.Count).End(xlUp).row
    For i = 10 To k
        Set initial_date = .Range("J2:AKS2").Find(Range("D" & i).Value)
        colum = Range(initial_date.Address).Column
        day = .Range("D" & i).Value
        
        Do While day <= .Range("E" & i).Value
            If .Cells(i, colum).Interior.Color = 16777215 Then
                .Cells(i, colum) = "MP"
                day = day + 1
                colum = colum + 1
            Else
                day = day + 1
                colum = colum + 1
            End If
        Loop
    Next i
End With



End Sub