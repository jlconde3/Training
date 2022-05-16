Sub PLAN_ESTUDIOS_AIT()

Dim FirstColumn As Integer, LastColumn As Integer, RowPlannig As Integer, ColumnPlanning As Integer, i As Integer, LastRowPlanning As Integer

Worksheets.Add.name = "TEMP"
RowPlanning = 10
i = 1

With Worksheets("PLANNING_FAM")
    LastRowPlanning = .Cells(.Rows.Count, "A").End(xlUp).row
    Do While RowPlanning <= LastRowPlanning
        If .Cells(RowPlanning, "I") = "FAM_AIT" Then
            For ColumnPlanning = 10 To 756
                If IsEmpty(.Cells(RowPlanning, ColumnPlanning)) = False Then
                    Worksheets("TEMP").Cells(i, 1) = .Cells(RowPlanning, "G")
                    Worksheets("TEMP").Cells(i, 2) = .Cells(2, ColumnPlanning)
                    i = i + 1
                End If
            Next ColumnPlanning
        End If
        RowPlanning = RowPlanning + 1
    Loop
    

End With


With Worksheets("TEMP")

    For j = 1 To i
        .Cells(j, 4) = .Cells(j, 1)
    Next j
    .Range("D1:D10000").RemoveDuplicates Columns:=Array(1), Header:=xlNo
    LastID = .Cells(.Rows.Count, "D").End(xlUp).row
    For j = 1 To LastID
        cont = 0
            For k = 1 To i
                If .Cells(j, 4) = .Cells(k, 1) Then
                    cont = cont + 1
                End If
            Next k
        .Cells(j, 5) = cont
    Next j
End With



Worksheets.Add.name = "PLAN_DE_ESTUDIOS"

'' Títulos de la hoja

With Worksheets("PLAN_DE_ESTUDIOS")

.Cells(1, 1) = "BILLET ID"
.Cells(1, 2) = "DATE"
.Cells(1, 3) = "CLASS 1"
.Cells(1, 4) = "CLASS 2"
.Cells(1, 5) = "CLASS 3"
.Cells(1, 6) = "CALSS 4"
.Cells(1, 7) = "CALSS 5"
.Cells(1, 8) = "CLASS 6"
.Cells(1, 9) = "MARK 1"
.Cells(1, 10) = "MARK 2"
.Cells(1, 11) = "MARK 3"
.Cells(1, 12) = "MARK 4"
.Cells(1, 13) = "MARK 5"
.Cells(1, 14) = "MARK 6"
End With

With Worksheets("TEMP")
LastRowPlanning = .Cells(.Rows.Count, "A").End(xlUp).row
j = 2
    For i = 1 To LastID
''--------------------------------------------------------------------------------------
        If .Cells(i, 5) >= 20 Then
            iData = 2
        End If
        
        If .Cells(i, 5) = "19" Then
            iData = 23
        End If
        
        If .Cells(i, 5) = "18" Then
            iData = 43
        End If
        
        If .Cells(i, 5) = "17" Then
            iData = 62
        End If
        
        If .Cells(i, 5) = "16" Then
            iData = 80
        End If
        
        If .Cells(i, 5) = "15" Then
            iData = 97
        End If
        
        If .Cells(i, 5) = "14" Then
            iData = 113
        End If
        
        If .Cells(i, 5) = "13" Then
            iData = 128
        End If
        
        If .Cells(i, 5) = "12" Then
            iData = 142
        End If
        
        If .Cells(i, 5) = "11" Then
            iData = 155
        End If
        
        If .Cells(i, 5) = "10" Then
            iData = 167
        End If
        
        If .Cells(i, 5) = "9" Then
            iData = 178
        End If
        
        If .Cells(i, 5) = "8" Then
            iData = 188
        End If
        
        If .Cells(i, 5) = "7" Then
            iData = 197
        End If
        
        If .Cells(i, 5) = "6" Then
            iData = 205
        End If
        
        If .Cells(i, 5) = "5" Then
            iData = 212
        End If
''------------------------------------------------------------------------------------------------
        For k = 1 To LastRowPlanning
            If .Cells(i, 4) = .Cells(k, 1) And IsEmpty(iData) = False Then
            Sheets("PLAN_DE_ESTUDIOS").Cells(j, 1) = .Cells(i, 4)
            Sheets("PLAN_DE_ESTUDIOS").Cells(j, 2) = .Cells(k, 2)
            Sheets("PLAN_DE_ESTUDIOS").Cells(j, 3) = Sheets("DATA").Cells(iData, 2)
            Sheets("PLAN_DE_ESTUDIOS").Cells(j, 4) = Sheets("DATA").Cells(iData, 3)
            Sheets("PLAN_DE_ESTUDIOS").Cells(j, 5) = Sheets("DATA").Cells(iData, 4)
            Sheets("PLAN_DE_ESTUDIOS").Cells(j, 6) = Sheets("DATA").Cells(iData, 5)
            Sheets("PLAN_DE_ESTUDIOS").Cells(j, 7) = Sheets("DATA").Cells(iData, 6)
            Sheets("PLAN_DE_ESTUDIOS").Cells(j, 8) = Sheets("DATA").Cells(iData, 7)
            iData = iData + 1
            j = j + 1
            End If
        Next k
''-----------------------------------------------------------------------------------------------
Next i
    
End With
 
Worksheets("TEMP").Delete


End Sub
